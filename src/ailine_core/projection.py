"""投影法 ── 表を別の形へ写すとき、**何を保存し、何を保存しないか**を宣言する。

★★ 出所（2026-09-02）: FORMAT_MAP の困難を ailine の語彙だけで書き、盲で立てた
  23 学問分野へ撃って拾った。★ 予測に入れていなかった分野から来た。

    地図の投影法 ── 面積・角度・距離のどれを保存するかを「選ぶ」。
                  全部同時には不可能（**定理**）。
                  だから投影法には名前がつき、地図には必ず明記される。

★ そこで読み替えが起きた ── 「依頼文に情報が無い」（依頼者側の欠陥）ではなく
  **「保存する不変量を選ばせていない」**（こちらの設計の穴）。

★★ なぜ効くか（2026-09-05 に実測で確かめた）:
  式のある表に様式写像を掛けると **必ず × になっていた**。元の `金額` が `=B2*C2`、
  出力は `36000` ── 事後条件は式の文字列と比べて「不一致」と判定していた。
  **正当な変換を事故として扱っていた。**
  ★ 欠けていたのは「この写像は**式を保存せず値を保存する**」という宣言。
    判定に要る三項（依頼・宣言・実体）のうち、また宣言だった。

★ この表は**新しい検算を作らない**。既に在る事後条件に**名前を付ける**だけ。
  そのうえで「保存しないもの」を書く前に人へ見せる（地図に投影法名を書くのと同じ）。

★ 手で列挙しない方針との関係: ここは **op ごとに違う写像の性質**そのものなので、
  宣言表として持つのが正しい（`OP_WRITE_TARGET` と同じ立場）。
  ★ ただし**漏れは番人が数える** ── 新しい derive op が増えたら赤くなる。
"""
from __future__ import annotations

# --- 保存されうるもの（不変量の名前）-------------------------------------------
VALUES = "値"
COLUMN_NAMES = "列名"
COLUMN_ORDER = "列の順序"
ROW_IDENTITY = "行の同一性"
ROW_COUNT = "行数"
FORMULAS = "式"
FORMATTING = "書式"
TOTAL_ROW_MEANING = "合計行の意味"
SOURCE_SHEET = "元のシート"

#: すべての不変量（番人が「宣言し忘れ」を数えるための全体集合）
ALL_INVARIANTS = (VALUES, COLUMN_NAMES, COLUMN_ORDER, ROW_IDENTITY, ROW_COUNT,
                  FORMULAS, FORMATTING, TOTAL_ROW_MEANING, SOURCE_SHEET)


class Projection:
    """1 つの op が宣言する写像の性質。

    keeps: 保存すると約束するもの（事後条件が確かめている）
    drops: **保存しないと明記する**もの（★ 投影法の値打ちはこちら側にある）
    why:   保存できない理由を人の言葉で（地図が「面積が歪む」と言うのと同じ）
    """

    __slots__ = ("keeps", "drops", "why")

    def __init__(self, keeps: tuple, drops: tuple, why: dict | None = None):
        self.keeps = tuple(keeps)
        self.drops = tuple(drops)
        self.why = dict(why or {})


#: 新しい表を作る 7 op（derive 群）の宣言。
#: ★ FORMAT_MAP だけに入れない ── 1 つの検体だけで測って一般化した過去の失敗を繰り返さない。
PROJECTIONS = {
    # 集計: 元の行は消え、グループごとの 1 行になる。値は足し合わされる。
    "AGGREGATE": Projection(
        keeps=(COLUMN_NAMES, TOTAL_ROW_MEANING, SOURCE_SHEET),
        drops=(ROW_IDENTITY, ROW_COUNT, FORMULAS, FORMATTING),
        why={ROW_IDENTITY: "グループごとに 1 行へまとめるため、元の行は残りません",
             FORMULAS: "集計結果は値として書きます"}),
    # ピボット: DataPilot。開き直すたび書式が消える癖がある（語彙表に明記済み）。
    "PIVOT": Projection(
        keeps=(COLUMN_NAMES, SOURCE_SHEET),
        drops=(ROW_IDENTITY, ROW_COUNT, FORMULAS, FORMATTING, TOTAL_ROW_MEANING),
        why={FORMATTING: "ピボットは開き直すたびに書式が消えます"
                          "（書式つきの見栄えが要るなら『集計』の方が向きます）",
             FORMULAS: "集計結果は値として書きます",
             TOTAL_ROW_MEANING: "ピボットが自分で総計行を作るので、"
                                 "元の合計行はそのまま持ち込みません"}),
    # 抽出: 条件に合う行だけを新しいシートへ。行はそのまま運ぶ。
    "EXTRACT": Projection(
        keeps=(VALUES, COLUMN_NAMES, COLUMN_ORDER, ROW_IDENTITY, SOURCE_SHEET),
        drops=(ROW_COUNT, FORMULAS, FORMATTING),
        why={FORMULAS: "抜き出した先には元の列が無いので、式は値に落とします"}),
    # 列抽出: 選んだ列だけを新しいシートへ。行は全部運ぶ。
    "EXTRACT_COLUMNS": Projection(
        keeps=(VALUES, COLUMN_NAMES, ROW_IDENTITY, ROW_COUNT, SOURCE_SHEET),
        drops=(COLUMN_ORDER, FORMULAS, FORMATTING),
        why={COLUMN_ORDER: "頼まれた列だけを、頼まれた並びで出します",
             FORMULAS: "抜き出した先には元の列が無いので、式は値に落とします"}),
    # 帳票: 1 行を雛形 1 枚へ。雛形の書式は残る（雛形そのものを複製するため）。
    "REPORT_PER_ROW": Projection(
        keeps=(VALUES, ROW_IDENTITY, FORMATTING, SOURCE_SHEET),
        drops=(COLUMN_NAMES, COLUMN_ORDER, ROW_COUNT, FORMULAS),
        why={COLUMN_NAMES: "雛形の見出しに置き換わります（印 {{列名}} の位置へ）",
             FORMULAS: "写した先には元の列が無いので、式は値に落とします"}),
    # 様式写像: 各行を雛形 1 行へ。N 行の新しい表になる。
    "FORMAT_MAP": Projection(
        keeps=(VALUES, ROW_IDENTITY, ROW_COUNT, SOURCE_SHEET),
        drops=(COLUMN_NAMES, COLUMN_ORDER, FORMULAS, FORMATTING),
        why={COLUMN_NAMES: "雛形の見出しに置き換わります（印 {{列名}} の位置へ）",
             FORMULAS: "写した先には元の列が無いので、式は値に落とします"}),
    # 参照埋め: その場で列を埋める。新しいシートは作らない。
    "LOOKUP_FILL": Projection(
        keeps=(VALUES, COLUMN_NAMES, COLUMN_ORDER, ROW_IDENTITY, ROW_COUNT,
               FORMATTING, TOTAL_ROW_MEANING),
        drops=(FORMULAS,),
        why={FORMULAS: "引いてきた値は値として書きます（対応表への参照は残しません）"}),
}


def projection_for(op: str | None):
    """その op の宣言（無ければ None ── 写す op でない）。"""
    return PROJECTIONS.get(op or "")


def render_projection_notice(op: str | None) -> list:
    """書く**前**に見せる 1〜2 行（地図に投影法名を書くのと同じ）。

    ★ 「保存する」より「**保存しない**」を先に言う ── そちらが人の判断を変える。
    """
    proj = projection_for(op)
    if proj is None or not proj.drops:
        return []
    out = [f"　　この写し方で保存されないもの: {'・'.join(proj.drops)}"]
    for name in proj.drops:
        why = proj.why.get(name)
        if why:
            out.append(f"　　　・{name}: {why}")
    return out
