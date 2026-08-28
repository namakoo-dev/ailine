"""report_group — 帳票段 REPORT_PER_GROUP の純ロジック部品。

★★ なぜ要るか（2026-08-28・Namakoo の指摘）:
  「同名の取引先から複数の発注があるケースでは請求書を一枚にまとめないといけない」
  REPORT_PER_ROW は **1 データ行 = 1 枚**。同じ取引先が 2 行あると請求書が 2 枚になる。
  契約としては ✓ が出る（宣言どおりに 2 枚作った）が、**仕事としては間違い**。
  ★ この repo の型そのもの: 宣言と実体は合っていて、**依頼（人が本当に欲しいもの）**が
    見られていない。だから直しは「宣言を正す」── 1 グループ = 1 枚に変える。

★ 憲法の適用は REPORT_PER_ROW と同じ: 雛形は人が作る。機械は印の在るセルだけ埋める。

★ 印は 3 種類。**どれも人が雛形に書く**（依頼文に新しい言い方を覚えさせない ──
  雛形が形を決めるので、一段目の語彙（OPS_DOC）は 1 文字も増えない）:
    {{列名}}          … その取引先ぜんぶで**同じはず**の値（取引先名・締め日・担当）
                        グループ内で食い違ったら**埋めずに断る**（どれを書けばいいか
                        機械には決められない ── 推測で 1 つ選ぶのが一番こわい）
    {{明細:列名}}     … 発注 1 件ごとの値。この印の在る行が**明細行**で、件数ぶん増える
    {{合計:列名}}     … そのグループの合計（数値列のみ・機械が足す）

★ ailine を import しない（report_per_row.py と同じ移植可能性の作法）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

# 印の接頭辞。**日本語の語**を使う（雛形を書くのは人であって機械ではない）。
DETAIL_PREFIX = "明細:"
TOTAL_PREFIX = "合計:"

_NUM_RE = re.compile(r"^-?[0-9]+(\.[0-9]+)?$")


@dataclass(frozen=True)
class GroupPlan:
    """1 グループ分の宣言。name: グループ名（＝シート名の元）。rows: 元の行番号（1起点・昇順）。"""
    name: str
    rows: tuple


@dataclass(frozen=True)
class MarkLayout:
    """雛形の印を 3 種類に仕分けた結果。
       detail_row: 明細行（1起点）。明細の印が 1 つも無ければ None。
       detail/total/value: それぞれの Placeholder の並び（走査順）。"""
    detail_row: int | None
    detail: tuple
    total: tuple
    value: tuple


def mark_kind(column_name: str) -> tuple:
    """印の中身 → (種類, 列名)。種類は "detail" / "total" / "value"。

    ★ 前後の空白は落とす（人が書くので『{{明細: 項目}}』は普通に起きる）。
    ★★ 全角コロンも半角と同じに読む: 日本語 IME で雛形を書けば『明細：項目』が出る方が
      自然で、半角しか見ないと「列『明細：項目』が見つかりません」という**筋違いの断り**
      になる（断り文が原因を指さないのが一番たちが悪い）。"""
    s = str(column_name).replace("：", ":")
    if s.startswith(DETAIL_PREFIX):
        return "detail", s[len(DETAIL_PREFIX):].strip()
    if s.startswith(TOTAL_PREFIX):
        return "total", s[len(TOTAL_PREFIX):].strip()
    return "value", s.strip()


def classify_placeholders(placeholders) -> tuple:
    """印の並び → (MarkLayout, 断りの文 or None)。

    ★ 断る条件は 2 つだけ（どちらも「埋めたら静かに壊れる」形）:
      ① 明細の印が **2 行以上**に散っている ── 何行ぶん増やせばいいか決まらない
      ② 明細/合計の印の列名が空（『{{明細:}}』のような書き間違い）
    """
    detail, total, value = [], [], []
    for ph in placeholders:
        kind, name = mark_kind(ph.column_name)
        if kind in ("detail", "total") and not name:
            return None, (f"雛形の印『{{{{{ph.column_name}}}}}』（{ph.cell}）に列名がありません。"
                           f"『{{{{明細:項目}}}}』のように列名まで書いてください")
        (detail if kind == "detail" else total if kind == "total" else value).append(ph)
    rows = sorted({ph.row for ph in detail})
    if len(rows) > 1:
        return None, (f"明細の印が {len(rows)} 行（{rows}）に散っています。"
                       "明細行は 1 行にまとめてください（その 1 行が件数ぶん増えます）")
    # ★★ 2026-08-28（設計査読で名指しされた・自分では見えていなかった穴）:
    #   明細行に {{列名}} や {{合計:}} が**同居**すると、その値が件数ぶん刷られる。
    #   埋める側と確かめる側が同じずれ関数を共有するので**事後条件は通り、✓ が出る**
    #   ── 宣言と実体は合っていて依頼だけが違う、この repo の三項の型そのもの。
    if rows:
        squatters = [ph for ph in list(total) + list(value) if ph.row == rows[0]]
        if squatters:
            names = "・".join(f"『{{{{{ph.column_name}}}}}』（{ph.cell}）" for ph in squatters[:3])
            return None, (f"明細行（{rows[0]}行目）に、明細でない印が同居しています: {names}。"
                           "明細行は件数ぶん増えるので、その印も件数ぶん刷られます ── "
                           "別の行へ移してください")
    return MarkLayout(detail_row=rows[0] if rows else None,
                      detail=tuple(detail), total=tuple(total), value=tuple(value)), None


def build_groups(rows: list, name_col_idx: int) -> list:
    """[(行番号, [セル値, ...]), ...] → GroupPlan の並び（**最初に出た順**）。

    ★ 並べ替えない: 人が表に並べた順が意味を持つことがある（並べ替えは別の操作）。
    ★ 名前が空の行は呼び出し側で除いてから渡す（ここでは判断しない）。
    """
    order, by_name = [], {}
    for row_no, vals in rows:
        name = "" if name_col_idx - 1 >= len(vals) else vals[name_col_idx - 1]
        key = "" if name is None else str(name).strip()
        if key not in by_name:
            by_name[key] = []
            order.append(key)
        by_name[key].append(row_no)
    return [GroupPlan(name=k, rows=tuple(by_name[k])) for k in order]


def value_conflicts(group: GroupPlan, values_by_row: dict, column_name: str) -> list:
    """グループ内で {{列名}} の値が食い違っていないか。食い違う値の一覧（0 か 2 つ以上）。

    ★ **食い違ったら埋めない**。1 枚の紙に 1 つしか書けない欄に、どちらを書くかは
      機械には決められない ── 推測で選ぶと、間違った担当者名が顧客に届く。
      合計したいなら人が『{{合計:列名}}』と書く（意図が印に出る）。
    """
    seen = []
    for row_no in group.rows:
        v = (values_by_row.get(row_no) or {}).get(column_name)
        if v not in seen:
            seen.append(v)
    return seen if len(seen) > 1 else []


def is_numeric(v) -> bool:
    """合計に足してよい値か。★ 文字列の '1,200' は足さない（表記を推測しない）。"""
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    return bool(isinstance(v, str) and _NUM_RE.match(v.strip()))


def sum_for(group: GroupPlan, values_by_row: dict, column_name: str) -> tuple:
    """(合計, 断りの文 or None)。数値でない値が 1 つでも混ざれば足さずに断る。

    ★ 全部が整数なら**整数で返す**。float を返すと、型込み等値で確かめる事後条件が
      30000 と 30000.0 の食い違いで偽の × を出す（設計査読の指摘）。
    """
    total = 0.0
    all_int = True
    for row_no in group.rows:
        v = (values_by_row.get(row_no) or {}).get(column_name)
        if v is None or v == "":
            continue
        if not is_numeric(v):
            return None, (f"『{column_name}』の {row_no}行目が数値ではありません（{v!r}）。"
                           "合計は数値の列にだけ置けます")
        f = float(v)
        if f != int(f):
            all_int = False
        total += f
    return (int(total) if all_int and total == int(total) else total), None


def needs_grouping(groups: list) -> bool:
    """1 つでも 2 行以上のグループが在るか（＝まとめないと紙が 2 枚になる）。"""
    return any(len(g.rows) > 1 for g in groups)


def output_rows_for(template_row: int, detail_row: int | None, n: int) -> list:
    """雛形の 1 行 → まとめた紙の上でどの行になるか（1起点）。

    ★ 明細行を n 行に増やすと、その**下は全部ずれる**。ずれを 1 箇所で決めておかないと、
      埋める側と確かめる側が別々にずれを数えることになり、片方だけ直る（この repo で
      何度も出た形）。埋める側は Basic だが、確かめる側はここを使う。
    """
    if detail_row is None or n <= 0:
        return [template_row]
    if template_row < detail_row:
        return [template_row]
    if template_row == detail_row:
        return [detail_row + k for k in range(n)]
    return [template_row + n - 1]


def detail_index_for(out_row: int, detail_row: int | None, n: int) -> int:
    """まとめた紙の行 → 何件目の明細か（0起点）。明細行でなければ 0。"""
    if detail_row is None:
        return 0
    k = out_row - detail_row
    return k if 0 <= k < n else 0
