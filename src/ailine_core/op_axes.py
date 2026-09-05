"""op の**軸**と、その軸の上でまだ持っていない値を宣言する（2026-09-05）。

★★ 出所: 断りの導線を 15 件で測ったら、提案の 3/9 が**意図とずれていた**。

    「単価の平均値を一番下に追加して」→ もしかして: 合計追加？
    （『平均値・一番下・追加』などの部分はこの操作に反映されません）

  ★ 道具は**自分でずれを分かっている**のに、提案として先頭に出していた。
    依頼の芯（平均）が「反映されません」に入っているなら、それは「近い操作」でなく
    **別の操作**だ。

★ Namakoo の指摘: 候補（提案の出し方を変える・既定を N にする等）は
  **どれも逃げで、問題を解いていない**。正しい応答は「**平均は扱えません**」と言うこと。

★★ 構造の穴: 知識は在るが**散文**だった。語彙表にはこう書いてある ──
    「★ 合計(SUM)専用。平均・最大・最小など他の統計量は語彙に無い」
  これは LLM への**指示**であって保証ではない（今日 4 回確認した形）。
  実際、判定器はこの文を読んでいるのに APPEND_TOTAL を出した。

★ 2026-08-22 に Namakoo が却下したのは **開集合**の名簿（道具が持たない全機能・
  増築しても収束しない）。ここで宣言するのは **閉集合** ── その op の軸の上の兄弟で、
  軸は op の意味が決めるので有限。性質が違う。

    APPEND_TOTAL   軸: 集約関数  持つ: 合計   まだ無い: 平均・最大・最小・件数・中央値
    COMPUTE_COLUMN 軸: 演算子    持つ: + - * / まだ無い: 累乗・剰余
    DEDUP          軸: 残す側    持つ: 最初の1行 まだ無い: 最後だけ残す・全部消す

★★ 判定は**実表に聞く**（実装前に 13 検体で測って決めた形・誤り 0）:

    ok        その語が無い / 列名の一部でしかない / 否定形で打ち消されている
    lacks     ★ 軸の上にまだ無い機能を名指ししている → 断って、持っている方の例を出す
    ambiguous ★ 列名と操作語が**同居**している → 「判断しかねる」と言って言い直しを勧める

  ★ ambiguous は Namakoo の指示で足した。黙って通すより正直で、実装は同じ複雑さだった。
  ★ 単純な部分文字列一致では **5/6 で誤爆**した（「平均単価の合計」を断る等）──
    実装前に測って分かった。実表を見る形にして 13/13。

★★ 「まだ無い」であって「持たない」ではない（Namakoo・2026-09-05）
  平均などは後に扱う可能性がある。**扱えるようになったら has へ動かすだけ**で、
  この表以外は触らなくてよい形にしてある（発火条件つきの保留）。
"""
from __future__ import annotations

#: 否定形（「全部消さずに」＝ 全部消すことを頼んでいない）
_NEGATIONS = ("ずに", "ないで", "なく")


class Axis:
    """1 つの op の軸。

    name:  軸の名前（人に見せる。「集約関数」「演算子」）
    has:   いま持っている値（人に見せる。「合計」）
    lacks: {依頼文に現れる語: 人に見せる名前} ── ★ **まだ**持っていないもの
    since: いつ・なぜ保留にしたか（発火条件つきで残す）
    """

    __slots__ = ("name", "has", "lacks", "note")

    def __init__(self, name: str, has: tuple, lacks: dict, note: str = ""):
        self.name, self.has, self.lacks, self.note = name, tuple(has), dict(lacks), note


AXES = {
    "APPEND_TOTAL": Axis(
        name="集約関数", has=("合計",),
        lacks={"平均": "平均", "平均値": "平均", "最大": "最大", "最大値": "最大",
               "最小": "最小", "最小値": "最小", "件数": "件数", "中央値": "中央値"},
        note="★ 保留（2026-09-05・Namakoo）: 平均などは**後に扱う可能性がある**。"
             "扱えるようになったら has へ動かすだけでよい。"
             "発火条件: 実需で平均/最大/最小が来たら、または合計以外の集約が"
             "語彙に入ったとき。"),
    "COMPUTE_COLUMN": Axis(
        name="演算子", has=("足し算", "引き算", "掛け算", "割り算"),
        lacks={"累乗": "累乗", "べき乗": "累乗", "剰余": "剰余"},
        note="★ 保留: 累乗・剰余は実需が来たら検討する。"),
    "DEDUP": Axis(
        name="残す行", has=("最初の 1 行を残す",),
        lacks={"最後の1行": "最後の 1 行だけ残す", "全部消": "重複を全部消す"},
        note="★ 保留: 「重複を全部消す」は残す側の選択が要る（どちらも消える）。"
             "発火条件: 実需で来たら、選ばせる形（投影法と同じ）で設計する。"),
}


def axis_for(op: str | None):
    return AXES.get(op or "")


def judge_axis(op: str | None, task: str, headers=None) -> tuple:
    """(判定, 詳細) を返す。判定は "ok" / "lacks" / "ambiguous"。

    詳細は ("人に見せる名前", "依頼文の語", "同居している列名 or None")。
    ★ headers（対象シートの実在列名）を渡さないと誤爆する ── 実表に聞くのが要点。
    """
    axis = axis_for(op)
    if axis is None or not task:
        return "ok", None
    cols = [str(h) for h in (headers or []) if h]
    for word, shown in axis.lacks.items():
        if word not in task:
            continue
        # ★ その語を含む列名が、依頼文に**そのまま**現れているか
        named = [h for h in cols if word in h and h in task]
        if named:
            # ★ 列名を消した残りにまだ在るなら、操作語としても使われている＝決められない
            rest = task
            for h in named:
                rest = rest.replace(h, "")
            if word in rest:
                return "ambiguous", (shown, word, named[0])
            continue
        # ★ 否定形が直後に来るなら「しない」と言っている
        i = task.find(word)
        if any(n in task[i:i + 8] for n in _NEGATIONS):
            continue
        return "lacks", (shown, word, None)
    return "ok", None


def render_axis_refusal(op: str | None, detail, label: str | None = None) -> list:
    """「まだ扱えません」を人の言葉で。★ 持っている方の言い方は examples.py が出す。"""
    axis = axis_for(op)
    if axis is None or not detail:
        return []
    shown, _word, _col = detail
    have = "・".join(axis.has)
    return [f"？ {shown}はまだ扱えません（この道具の{axis.name}は {have} だけです）"]


def render_axis_ambiguity(op: str | None, detail) -> list:
    """「判断しかねる」を人の言葉で（★ Namakoo の指示・2026-09-05）。

    ★ 黙って通すより正直 ── 列名と操作語が同居していると、機械には決められない。
    """
    if not detail:
        return []
    shown, word, col = detail
    return [f"？ 『{word}』が列名（{col}）を指しているのか、{shown}という操作を"
            "指しているのか判断しかねます",
            f"  どちらか分かる言い方に直してください"
            f"（列として使うなら「{col}の合計を一番下に追加して」のように）"]
