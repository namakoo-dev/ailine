"""事後条件の層が**群をまたいで共有する**もの ── 読み方と、開示の言い方。

★ なぜ在るか（2026-09-03・分割の二歩目）: op ごとの事後条件を
shape / move / derive の 3 群に割るとき、**2 群以上から呼ばれる 8 本**がある。
どれか 1 群に置くと、他の群がその群を import することになり、
「並べ替えの検算が、なぜか書式の群に依存している」という読めない形になる。

★ ここに置く線: **どの群から呼ばれても意味が変わらないもの**だけ。
  ・読み方 … `_numeric_value` / `_row_as_shown` / `_cells_for_shift` / `_extract_predicate`
  ・突き合わせ … `compare_moved_rows`（★ 挿入も削除もここ 1 箇所を通す）
  ・開示の言い方 … `note_unverified` / `note_stringy_numbers` / `_moved_rows_note`

★ 「開示の言い方」を共有する理由: 同じ事情を op ごとに違う文で言うと、
**同じ穴なのに別の穴に見える**。この repo は「判定は正しいが説明の文面が古い」で
1 日潰したことがある（2026-08-31）。文を 1 箇所に持つのは、その再発を止めるため。

★ `compare_moved_rows` が単独で在る意味: 行が「動いただけ」かを見る検算は、
挿入・削除・並べ替え・入れ替えのすべてが要る。**片方だけ直す**のを防ぐために
1 関数へ畳んである（系譜「二重化した経路は片配線が既定で起きる」）。
"""
from __future__ import annotations

from datetime import date as _date_cls, datetime

from ailine_core.date_compare import date_to_serial
from ailine_core.primitives import is_number as _is_number

def _numeric_value(v):
    """セルの値を「表計算にとっての数値」にする。数値でなければ None。

    ★ なぜ在るか（2026-08-24 の実測）: 出納帳を「日付の古い順に並べ替えて」と頼むと、
      LibreOffice は正しく並べたのに事後条件が
      「検証対象が0件（数値でない 3 行は対象外）」で拒否して原本に反映しなかった。
      **表計算の日付はシリアル値という数値**であり、openpyxl が datetime を返すせいで
      検証側だけがそれを見失っていた。日付を扱う道具で日付が並べ替えられないのは致命的。
    ★ 日付→シリアル値の換算は ailine_core.date_compare に 1 箇所だけ置く
      （EXTRACT の日付比較と同じ換算 ── 2 つ持つと片方だけ直る）。
    """
    if _is_number(v):
        return float(v)
    if isinstance(v, datetime):
        return float(date_to_serial(v.date())) + (
            v.hour * 3600 + v.minute * 60 + v.second) / 86400.0
    if isinstance(v, _date_cls):
        return float(date_to_serial(v))
    return None

def note_stringy_numbers(args: dict, values) -> None:
    """数え上げの対象に「数字に見える文字列」が混ざっていたら、機械の値として残す。

    ★ 2026-08-25（塊③・中核 op の致命3/4）: `check_aggregate` は
      `v = v if _is_number(v) else 0` を**期待側と観測側の両方**に掛けていた。
      LibreOffice の SUM と同じ落とし方をするので**必ず一致する** ── 検算が恒真。
      実測: 営業 1000 + '2000'(文字列) + 1500 → 出力 2500（正 4500）で ✓ が出た。
    ★ 恒真を切る条件は「**別実装で**確かめる」こと。文字列を 0 にする判断（_is_number）
      とは別の実装（compare_blocked.looks_numeric）で「数字に見えるか」を見る。
    ★ 判定（0 として足すこと）は 1 ビットも変えない ── 変えると LO の結果と食い違い、
      別の嘘になる。変えるのは**言うかどうか**だけ。
    """
    from ailine_core import compare_blocked as _cb
    stringy = [v for v in values if _cb.looks_numeric(v)]
    if not stringy:
        return
    shown = "、".join(f"『{v}』" for v in stringy[:3])
    more = f"（ほか {len(stringy) - 3} 件）" if len(stringy) > 3 else ""
    note_unverified(args, len(stringy),
                    f"数字に見えますが文字列のセルがあり、0 として足されています: "
                    f"{shown}{more} ── 合計がその分だけ小さくなります")

def note_unverified(args: dict, count: int, why: str) -> None:
    """検証できなかった行を、**機械の値として** args に残す。

    ★ 2026-08-25（塊①）: 3 面の盲検が別々の入口から着いた根 ──
      「判定に要る項が機械の値として在るのに、判定へ渡していない」。
      実測（中核 op の致命6）: check_sort は除外行数を int で数えたうえで、
      それを**文章にして捨て**、`pass` を返していた。結果、単価列の先頭が 250 なのに
      「✓ 機械検証済み」が出た ── 除外されたのは**まさに主張を壊す 2 行**だった。

    ★ ここでやらないこと: 判定（pass/fail）は 1 ビットも変えない。
      8 行を本当に検証したことは事実なので、⚠（機械保証なし）へは落とさない。
      ✓ を △ に降ろすのは決裁③の既存の機構（⚠ 始まりの行を数える）に任せる。

    ★ 表示文から読み取らせない: 文言を変えた瞬間に壊れる（この repo の既定の作法）。
    """
    if not isinstance(args, dict) or count <= 0:
        return
    args.setdefault("_unverified", []).append({"rows": int(count), "why": why})

def _cells_for_shift(bv, sheet_name, header_row: int, last_row: int, last_col: int) -> list:
    """行が**丸ごと動くだけ**のはずの操作（挿入・削除）で突き合わせるための読み。
       セルごとに (式か, 式ビューの値, 比べる値) を持つ ── 式セルの「比べる値」は
       **キャッシュ値**（計算後の値）。"""
    ws = bv.sheet(sheet_name)
    rows = []
    for r in range(header_row + 1, last_row + 1):
        row = []
        for c in range(1, last_col + 1):
            raw = ws.cell(row=r, column=c).value
            is_f = isinstance(raw, str) and raw.startswith("=")
            row.append((is_f, raw, bv.cell_value(r, c, sheet_name) if is_f else raw))
        rows.append(tuple(row))
    return rows

def compare_moved_rows(after_rows: list, before_rows: list, label: str) -> tuple:
    """**動いただけのはずの行**を突き合わせる ── 挿入も削除もここ 1 箇所を通す。

    ★ 2026-08-27 に実測して分かったこと（デモの経路で × が出た）:
      式は行が動くと**参照が自動で追随する**（`=B4-C4` が `=B5-C5` になる）。
      だから式を**文字で比べると必ず食い違い**、正しく押し下げた操作を
      「上書きした疑い」と誤って落とす。実際に落ちた。

    ★ かといって式セルを見ないことにはしない（見ない範囲は嘘の温床）。
      式セルは**キャッシュ値**（計算後の値）で見る:
        ・値が同じ  → その行は本当に「動いただけ」。確かめられた
        ・値が違う  → 落とさずに**開示する**。挿入した行を巻き込む合計式なら
                      正当に変わるが、機械にその区別はつかない（断れない時は開示する）
      式が消えた／増えたのは**壊れ**（落とす）。

    返り値: ("broken", 理由) または ("ok", 開示すべき食い違いの一覧)
    """
    if len(after_rows) != len(before_rows):
        return "broken", f"{label}: 行数が合いません"
    disclosures = []
    for i, (arow, brow) in enumerate(zip(after_rows, before_rows), start=1):
        for j, (a, bcell) in enumerate(zip(arow, brow), start=1):
            a_is_f, a_raw, a_val = a
            b_is_f, b_raw, b_val = bcell
            if a_is_f != b_is_f:
                return "broken", (f"{label}: {i} 行目の {j} 列目で式が"
                                   + ("消えました" if b_is_f else "増えました"))
            if not a_is_f:
                if a_raw != b_raw:
                    return "broken", f"{label}: {i} 行目の {j} 列目の値が変わっています"
                continue
            if a_val != b_val:
                disclosures.append(f"{i} 行目の {j} 列目（式の結果 {b_val!r}→{a_val!r}）")
    return "ok", disclosures

_MOVED_ROWS_WHY = ("動いた行の式の結果が変わっています ── 追加/削除した行を参照する"
                    "合計式なら正当ですが、機械には区別がつきません")

def _moved_rows_note(disclosures: list, why: str = _MOVED_ROWS_WHY) -> str:
    """開示の 1 行にまとめる（★ 決裁③: この文が付いた回は ✓ を出さない）。

    ★★ 2026-08-29（Namakoo「この操作が拒否されるのは正答か？」）: 断りは正しかった
      （実測: 列を入れ替えたら合計式が `=SUM(E2:INDEX(F:F,…))` と**二列にまたがり**、
       両方の合計が 1,000,440 ＝ 金額＋税込み金額 になっていた）。
      ★ ところが理由の文は**行の話**をしていた ── 列の入れ替えなのに
        「追加/削除した行を参照する合計式なら正当ですが」。
      ★ 正しい判定を、間違った言葉で説明していた。呼ぶ側が理由を渡せるようにする
        （文面を写し取らない ── 並べ方はここ 1 箇所のまま）。
    """
    head = "／".join(disclosures[:5])
    more = f"（ほか {len(disclosures) - 5} 件）" if len(disclosures) > 5 else ""
    return f"{why}: {head}{more}"

def _extract_predicate(cmp: str, threshold, date_mode: bool = False):
    """EXTRACT の判定を Basic 側(ExtractRows/helpers/AiLineHelpers.bas)とは別実装で
       もう一度書く（同じ勘定を2箇所が違う実装で書いて一致を見る・独立測定）。
       ★ M2（2026-08-21・宣言済みの挙動変更）: 意味論を tests/test_predicate_truth_table.py
       の手書きの表に合わせた。① eq は両辺が数値なら**許容誤差 1e-6**（浮動小数の完全一致は
       表計算の実データで偽陰性になる）② contains は**文字列セルのみ**（数値 140000 を
       黙って "140000" に文字列化して『40 を含む』としない ── 型の保存の哲学）。
       単一ブック EXTRACT（check_extract）の挙動もこの線に揃う。"""
    def _match(cell_value) -> bool:
        if cmp == "in":
            # ★ 丸ごと一致の集合判定（部分一致にしない ── 「りんご」が「青りんご」に
            #   当たると、頼んでいない行が黙って混じる）。空欄は一覧に入れない。
            if cell_value is None or cell_value == "":
                return False
            return str(cell_value) in {str(x) for x in (threshold or ())}
        if cmp == "nin":
            # ★★ 2026-09-02: 「〜以外」。**`in` の否定**として書く（別の勘定にしない）。
            #   ★ 空欄は「どれでもない」に**入れる** ── 「味噌汁以外を抜き出して」で
            #     名前が空の行を落とすと、残したい行を失う（取り返しのつかない側）。
            #   ★ 部分一致の否定にはしない（「青りんご」が「りんご以外」から外れる）。
            if cell_value is None or cell_value == "":
                return True
            return str(cell_value) not in {str(x) for x in (threshold or ())}
        if cmp == "contains":
            return (isinstance(cell_value, str) and threshold is not None
                    and str(threshold) in cell_value)
        if cmp == "eq":
            if isinstance(threshold, (int, float)) and not isinstance(threshold, bool):
                return (_is_number(cell_value)
                        and abs(float(cell_value) - float(threshold)) <= 1e-6)
            return str(cell_value) == str(threshold)
        # ★ 2026-08-24: 日付として比較するのは **閾値が日付リテラルだった時だけ**
        #   （date_mode）。凍結した真理値表が正しく縛っているとおり、素の数値 40000 に
        #   日付セル（シリアル値 46235）が当たってはいけない ── 「金額が40000以上」で
        #   日付列を選んでしまった時に全行一致する事故になる。
        v = _numeric_value(cell_value) if date_mode else (
            float(cell_value) if _is_number(cell_value) else None)
        if v is None:   # gte/lte/gt/lt は数値比較のみ
            return False
        t = float(threshold)
        if cmp == "gte":
            return v >= t
        if cmp == "lte":
            return v <= t
        if cmp == "gt":
            return v > t
        if cmp == "lt":
            return v < t
        return False
    return _match

def _row_as_shown(bv, sheet_name: str, row: int, last_col: int) -> list:
    """その行の「**見えている値**」（式のセルは計算結果）。

    ★★ 2026-08-30（Namakoo「特定条件の行や列の抜き出しができない」）:
      「丸和物流とみどり建設を抽出して」── **抽出そのものは成功していた**（7 行中 2 行）。
      落ちたのは検算で、元の `=E2*1.1`（式）と、抽出先の `63360`（値）を
      **文字どおり比べて**「元と不一致」と言っていた。
    ★ 抽出が値を写すのは正しい ── 式をそのまま持っていけば、新しいシートでは
      違うセルを指す。だから比べる相手は**計算結果**でなければならない
      （並べ替えの検算 compare_moved_rows が既に取っている線と同じ）。
    ★ 抽出と重複削除の**両方**が同じ形で比べていた ── 片方だけ直さない。
    """
    out = []
    for c in range(1, last_col + 1):
        f = bv.cell_formula(row, c, sheet_name)
        out.append(bv.cell_value(row, c, sheet_name) if f is not None
                    else bv.sheet(sheet_name).cell(row=row, column=c).value)
    return out
