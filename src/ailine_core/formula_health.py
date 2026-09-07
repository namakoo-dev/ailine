"""formula_health — 宣言つき挙動変更 #1: 型破壊の安全網。

★ ブラインド査定の実測: `原価` 列（数値）を SET_COLUMN_VALUE で文字列 `"0円"` に一括書換
すると、その列を参照する数式（例: `=D2-C2`）が `#VALUE!` に壊れる。なのに事後条件
チェッカー(check_set_column_value)は「対象列が指定文字列 N 件になったか」だけを見ており
「✓ 達成を機械検証済み」が出ていた（波及した数式破壊を誰も見ていない）。

(a) 一般の網: 適用の前後でエラー値セル(#VALUE!/#REF!/#DIV/0!/#N/A 等)を比較し、新たに
    エラーになったセルがあれば助言を返す。型破壊に限らずあらゆる波及被害を対象にする
    （査定者の言う「一番怖い種類の事故」への対抗・こちらが本命）。
(b) 狙い撃ちの関所……ではなく助言: OP_WRITE_TARGET が宣言する書き込み先列で、数値だった
    セルが数値に見えない文字列に変わった場合に予防的な確認行を返す。★ 関所（対話確認/
    非対話は停止）にはしていない ── 型が変わること自体は「税抜き列を税込みへ書き換える」
    のような正常な用途でも起き得るし、"0円" のように非数値文字列を意図的に書きたい
    場面もあるため、機械が意図まで断定できない（誤検知が出やすい・査定の+$30要望その2への
    予防止まりの応答という設計判断）。実害（数式が実際に壊れたか）は (a) が別途拾う。

★ ailine.py の行数は tests/ailine_py_line_budget.txt で凍結済み（増える方向の更新は禁止・
tests/test_line_budget.py 参照）。新しいコードは C4〜C7 の他モジュールと同じ理由でここに
置く。OP_WRITE_TARGET/`_is_number` 等 ailine.py 側の宣言・判定は、モジュール読み込み時点
ではなく呼び出し時点で引数として受け取る（dsl_step.DslStepDeps と同じ理由 ── テストの
monkeypatch より先に関数参照を固定しない）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Callable

import openpyxl
from openpyxl.utils import get_column_letter

from ailine_core.xml_readback import numeric_cells_became_strings

# OOXML のエラー値。openpyxl は data_only=True で開くと、LibreOffice/Excel が保存時に
# 計算・保存したキャッシュ値をそのまま文字列として返す（エラー値も同じ経路で読める）。
_ERROR_VALUES = frozenset({
    "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!",
})


def _error_cells(path: Path) -> dict:
    """path をキャッシュ値(data_only)で開き、エラー値を持つセルを
       {(sheet, row, col): エラー文字列, ...} で返す。式そのものは openpyxl が計算しない
       ため、ここで見えるのは basrun_apply（LibreOffice 経由で保存）を通した後のファイル
       だけ。ファイルが開けない/壊れている場合は空 dict（保守的＝誤検知回避）。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return {}
    out: dict = {}
    try:
        for name in wb.sheetnames:
            for row in wb[name].iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v in _ERROR_VALUES:
                        out[(name, cell.row, cell.column)] = v
    finally:
        wb.close()
    return out


def new_error_cells(before_path, after_path) -> dict:
    """適用で**新たに**エラーになったセル {(シート, 行, 列): 値}（無ければ空）。

    ★★ 2026-09-08（盲検の検品が拾った）: 「丸和物流の行を削除して」で合計式が
      `=SUM(#REF!:INDEX(E:E,ROW()-1))` に壊れ、**原本に `#REF!` が残った**。
      ★ 検出はこの器官が既にしていた（助言として ⚠ を出し ✓ も降ろしていた）──
        足りなかったのは**帰結**で、「言ったうえで書いて」いた。
      ★ だから計算はここ 1 箇所のまま、呼ぶ側（反映の関所）を増やす。

    ★ 総数でなく**セル単位の新規発生**で見る（総数が同じでも、あるセルが壊れ別のセルが
      直った場合は別の問題）。
    """
    before_errors = _error_cells(before_path)
    after_errors = _error_cells(after_path)
    return {k: v for k, v in after_errors.items() if k not in before_errors}


def formula_error_advisory(before_path: Path, after_path: Path, *, cell_ref: Callable) -> list:
    """(a) 適用前後でエラー値セルを比較し、新たにエラーになったセル（前はエラーでなかった・
       今はエラー）が1件でもあれば助言を返す。
       ★ 単純な総数比較ではなくセル単位の新規発生で見る ── 総数が同じでも「あるセルが
       壊れ・別のセルが直った」場合は実質は別の問題であり、総数増減だけの比較では見逃す
       （ブリーフの「個数を数え、増えていたら報告」の趣旨をセル単位に強めた設計判断）。
       cell_ref: (row, col) -> "B2" のような表示用文字列に変換する関数（ailine.py の
       `_cell_ref` を渡す想定・表示フォーマットを二重管理しない）。"""
    new_breaks = new_error_cells(before_path, after_path)
    if not new_breaks:
        return []
    items = sorted(new_breaks.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2]))
    shown = items[:5]
    refs = "、".join(f"{sheet}!{cell_ref(r, c)}={val}" for (sheet, r, c), val in shown)
    more = f"、ほか{len(items) - 5}件" if len(items) > 5 else ""
    return [f"★ 疑わしい: 適用後にエラー値のセルが増えました（計{len(items)}件）: {refs}{more}"]


def _formula_ratio_by_column(path, header_rows: dict | None = None) -> dict:
    """{(シート, 見出し): (式のセル数, データ行数)} を返す。

    ★★ 見出し行は**受け取る**（2026-09-06・実物で鳴らずに分かった）: 初版は 1 行目と
      決め打ちしていたが、実物の請求書は**見出しが 16 行目**で 1 行目は空だった ──
      全列が無名になり、名簿が空になって**黙っていた**。合成検体では鳴っていたので、
      実物に当てるまで気づけなかった（検体は治具まで含めて仮説）。
    ★ 呼び出し側は `book_meta["header_rows"]` を既に持っている ── 渡すだけ。
      渡されなければ 1 行目（従来どおり・単体で使う経路のため）。

    ★ 位置でなく**見出しの名前**で持つ ── 列が動いても追える（`row_identity` が
      2026-09-02 に同じ問題で「名前で対応づける」に直したのと同じ手）。
      同じ見出しが 2 本ある表は最初の 1 本だけ見る（決められないものを決めない）。
    ★ 読めない回は空を返す（黙る＝断定しない）。
    """
    try:
        wb = openpyxl.load_workbook(path)          # ★ 式ビュー（data_only=False）
    except Exception:
        return {}
    out: dict = {}
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            hr = int((header_rows or {}).get(name, 1) or 1)
            heads, seen = {}, set()
            for c in range(1, (ws.max_column or 0) + 1):
                h = ws.cell(row=hr, column=c).value
                key = str(h).strip() if h is not None else ""
                if key and key not in seen:
                    seen.add(key)
                    heads[c] = key
            for c, key in heads.items():
                cells = [ws.cell(row=r, column=c).value
                         for r in range(hr + 1, (ws.max_row or 1) + 1)]
                if not cells:
                    continue
                out[(name, key)] = (sum(1 for v in cells if _looks_like_formula(v)),
                                    len(cells))
    except Exception:
        return out
    finally:
        wb.close()
    return out


def formula_loss_advisory(before_path, after_path, header_rows: dict | None = None) -> list:
    """**式だったセルが値に潰された**ら言う（直さない）。

    ★★ なぜ要るか（2026-09-05 の測定・2026-09-06 実装）: 実物の請求書には式が 29 個
      あった（`=IF(I17="","",G17*I17)` が 21 個 ほか）。そこへ「金額を全部 0 にして」と
      頼むと **21 個の式が消えて、事後条件は pass** する ── 宣言（列を 0 にする）は
      真だからだ。だが表としては壊れている: 以後 数量 を直しても金額が動かない。
    ★ 2026-08-31 の事故（`row_identity` を生んだ「2 セルだけ入れ替えた」は真だが
      表が矛盾）と**同じ形**。処置も同じ ── **直さない・言う**。

    ★★ 何を数えるかを実測で決めた（総数では区別できない）:
        値で潰す   式 3/3 → 0/3   ★ 言う
        行を削除   式 3/3 → 2/2   黙る（割合は変わらない）
        行を追加   式 3/3 → 4/4   黙る（`FillFormulasFromNeighbour` が運ぶ）
      ★ **列ごとの「式だったセルの割合」**で見れば、行の増減・並べ替えに影響されない。
        op の宣言を見ないので、**op が増えても配線が要らない**。
    ★ 8/27 の実測「式の文字列は行が動けば変わるのが正しい」に従い、**式の中身は比べない**。

    ★ 黙る場合（誤検知を避ける・どれも「頼まれたとおり」か「決められない」）:
      ・前に無かった列 / 後に無くなった列（列の追加・削除）
      ・前の割合が 0（元から式が無い）
      ・後のデータ行が 0（割合が計算できない）
    """
    before = _formula_ratio_by_column(before_path, header_rows)
    after = _formula_ratio_by_column(after_path, header_rows)
    if not before or not after:
        return []
    lost = []
    for key, (bf, bt) in before.items():
        if key not in after or bt == 0 or bf == 0:
            continue
        af, at = after[key]
        if at == 0:
            continue
        if af / at < bf / bt:
            lost.append((key[1], bf, bt, af, at))
    if not lost:
        return []
    lost.sort(key=lambda x: (x[3] / x[4]) - (x[1] / x[2]))
    head = lost[0]
    more = f"、ほか{len(lost) - 1}列" if len(lost) > 1 else ""
    return [f"⚠ 列『{head[0]}』は式で計算されていましたが、"
            f"{head[1]}/{head[2]} 件あった式が {head[3]}/{head[4]} 件になりました{more} ── "
            "以後この列は元の列を直しても追随しません（直していません）"]


def before_after_advisories(before_path, after_path, resolved: dict | None, *,
                            cell_ref, identity_advisory=None,
                            header_rows: dict | None = None) -> list:
    """**適用の前後を見比べて言う助言**を、ここ 1 本にまとめる。

    ★★ なぜ 1 本か（2026-09-06）: これまで呼び出し側が
      `formula_error_advisory(...) + broken_identity_advisory(...)` と**手で並べて**
      いた ── しかも同じ並びが **4 箇所**（単発・確認つき・上書き・複合計画の段）に
      写経されていた。3 本目を足せば 4 箇所 × 3 本になる。
      昨日 1 日かけて潰した「二重化した経路は片配線が既定で起きる」を、自分で
      1 世代増やすところだった。
    ★ 足す時はこの中に 1 行 ── 呼び出し側は 1 文字も変えない。
      （実演: 2026-09-06 に `formula_loss_advisory` を足したとき、5 つの呼び出し側は
        1 文字も変えていない ── 畳んでおいた甲斐が出た所）
    ★ 引数で受けるのは ailine.py 側の関数（`_cell_ref` / `broken_identity_advisory`）。
      モジュール読み込み時に固定しないのは、この repo の作法（テストの monkeypatch が
      効くようにする・dsl_step.DslStepDeps と同じ理由）。

    ★★ **宣言が要る助言と、要らない助言を分ける**（2026-09-06・番人に教わった）:
      自由生成の段は **op が決まっていないので `resolved`（宣言）が無い**。
      そこへ「等式が崩れた」を配ると、比べる相手が無いまま鳴る（過去に 5 箇所すべてへ
      配って NameError を 2 回出し、意図的に外した経緯が
      `tests/test_row_identity.py` に凍結されている）。
      ★ だから **入口は 1 本のまま、中で分ける** ── `identity_advisory=None` は
        「この経路には宣言が無い」の意味で、呼び出し側が渡さないことで表明する。
      ★ 「同じ形が 2 箇所にあれば片配線」ではない ── **意図的な非対称**が在る。
        非対称の理由は、こうしてコードと番人の両方に書いておく。
    """
    out = (formula_error_advisory(before_path, after_path, cell_ref=cell_ref)
           + formula_loss_advisory(before_path, after_path, header_rows))
    if identity_advisory is not None:
        out = out + identity_advisory(before_path, after_path,
                                      resolved if isinstance(resolved, dict) else {})
    return out


def _parses_as_number(s: str) -> bool:
    """文字列が数値そのものの見た目か（桁区切りのカンマは許容）。"105" は数値扱い・
       "0円" は非数値扱い ── SET_COLUMN_VALUE は常に setString で書くため、書いた中身が
       数字そのものか単位付き等の非数値文字列かで型変化の実害の芽の有無を分ける。"""
    try:
        float(s.replace(",", ""))
        return True
    except (ValueError, TypeError, AttributeError):
        return False


def _looks_like_formula(s) -> bool:
    """数式セルの見た目（"=B2*C2" 等）か。openpyxl の素の .value（data_only=False）は
       数式セルなら常にこの形の文字列を返す ── それ自体は「型が壊れた」証拠にならない
       （★ operator10 ⑤）。"""
    return isinstance(s, str) and s.strip().startswith("=")


def detect_write_target_type_change(before: dict, after: dict, *, op: str | None, resolved: dict | None,
                                     meta: dict | None, op_write_target: dict, is_number: Callable,
                                     after_path=None) -> str | None:
    """(b) OP_WRITE_TARGET が宣言する書き込み先列（狙い撃ち）で、変化前が数値・変化後が
       数値に見えない文字列のセルが1件でもあれば予防の確認行を返す（関所にはしない・
       理由はモジュール docstring 参照）。新規列（元の型という概念が無い）は対象外。
       ★ 単位C: op_write_target の値は `.col_key` / `.sheet_key` を持つ宣言オブジェクト
       （呼び出し側が渡す・ここは属性を読むだけで ailine を import しない）。
       ★ operator10 ⑤: before/after は snapshot()（openpyxl の素の .value・data_only=False）
       由来のため、数式セル（例 `=B2*C2`）は after 側の値が数式文字列そのものになり、
       キャッシュ値が数値でも「文字列に変わった」と誤検出していた（偽の破壊アラーム）。
       数式の見た目をした値は、after_path（適用後の実ファイル）が渡されていれば
       numeric_cells_became_strings でキャッシュ値を読み直して判定し、無ければ
       安全側（数式の文字列表現そのものは型変化の証拠にしない＝カウントしない）に倒す。"""
    if not (op and resolved is not None and meta is not None):
        return None
    write_target = op_write_target.get(op)
    if not write_target or not write_target.col_key:
        return None
    col_key, sheet_key = write_target.col_key, write_target.sheet_key
    col_name = resolved.get(col_key)
    if not col_name:
        return None
    sheet = resolved.get(sheet_key) if sheet_key else next(iter(meta.get("sheets") or []), None)
    if not sheet or sheet not in before.get("sheets", []):
        return None
    headers = meta.get("headers", {}).get(sheet, [])
    if col_name not in headers:
        return None
    col_idx = headers.index(col_name) + 1
    prefix = sheet + "!"
    flips = 0
    for k, b in before["cells"].items():
        if not k.startswith(prefix):
            continue
        _, rc = k.split("!", 1)
        r_str, c_str = rc.split(",")
        if int(c_str) != col_idx:
            continue
        if not is_number(b[0] if b else None):
            continue
        a = after["cells"].get(k)
        a_val = a[0] if a else None
        if not (isinstance(a_val, str) and not _parses_as_number(a_val)):
            continue
        if _looks_like_formula(a_val):
            if after_path is not None:
                ref = f"{get_column_letter(int(c_str))}{r_str}"
                if numeric_cells_became_strings(after_path, [ref], sheet_name=sheet):
                    flips += 1
            continue
        flips += 1
    if flips == 0:
        return None
    return (f"（確認）列『{col_name}』は元は数値でしたが、{flips} 件のセルが数値に見えない"
            f"文字列に変わりました。この列を参照する数式があれば壊れていないか確認してください")
