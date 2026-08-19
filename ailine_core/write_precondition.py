"""write_precondition — 単位F: 宣言した書き込み先の「前提」を、反映の直前に確かめる。

★ 盲検査定の実測（致命2件）:
  - 既存の合計行（A 列が空）があるブックへ「合計を出して」 → その行の値 116600 を
    `=SUM(...)` で**上書き**して 106000 になった。画面には `D4: 値 116600→'=SUM(...)'`
    と出ているのに、破壊の関所は鳴らず `✓ 機械検証済み` が出て exit 0。
  - 無関係な手作りの『集計』シート（年度/予算）があるブックへ「顧客ごとに集計して」 →
    SummaryTable ヘルパの `If oDoc.Sheets.hasByName("集計") Then removeByName("集計")`
    により中身が全滅。y/N も無く exit 0、しかも
    「（既存シート『集計』の更新は意図どおりです）」という**肯定文**まで出ていた。

★ 真因（独立レビュー）: 破壊の関所が「列」しか守っていなかった。関所の入口は
  `write_target.col_key` が無ければ即 return で、col_key を持つのは 16 op 中 3 つだけ。
  行に書く op も、シートに書く op も、守る腕が無かった。

★ 設計: **`writes` が宣言する領域には、成立していなければならない前提がある。
  関所はその前提が破れた時に鳴る。**

  | writes           | 前提                                   |
  |------------------|----------------------------------------|
  | new_column       | その列は空（＝既存の関所が守る）        |
  | new_row_at_end   | 書き込んだ行は before で空だった         |
  | new_sheet        | その名前のシートは before に存在しない   |
  | existing_column  | （前提なし＝上書き前提・既存の関所）     |
  | format_only      | 値が1つも変わらない                     |
  | row_shift/reorder| 値の多重集合が保存される                 |

★ 判定は「適用前の予測」でなく「**適用後の実測**」で行える ── 全工程は out_book（原本の
  コピー）の上で走り、原本への反映は最後の1手（atomic_replace_inplace）。before/after が
  両方手元にあって原本はまだ無傷、という窓がこの検査の置き場所。Basic のコード生成には
  1行も触らない。

★ 関所そのものはここには無い: この関数が返すのは**1行の文言**だけで、止める/聞く判断は
  呼び出し側の `_confirm_overwrite_or_gate`（既存の破壊の関所）がそのまま行う。
  新しい関所も新しい exit code も作らない ── 関所は1種類のまま、鳴る理由が増えるだけ。

★ 置き場所: ailine_core/（formula_health.py と同じ理由）。ailine.py は import しない ──
  宣言（OP_WRITE_TARGET の writes）も表示用の関数も、呼び出し**時点**で引数として受け取る。
"""
from __future__ import annotations

from collections import Counter
from typing import Callable

# 表示する実例の上限（多すぎると1行が読めなくなる。件数は必ず全数を述べる）。
_MAX_SAMPLES = 3


def _values(snap: dict) -> dict:
    """snapshot() の cells から「値のあるセル」だけを {(シート, 行, 列): 値} で取り出す。

    ★ 値だけを見る（書式・罫線・列幅は無視）。format_only の前提が「値が1つも変わらない」
    である以上、比較の土俵は値でなければならないし、LibreOffice 往復で書式表現が微妙に
    変わることは日常的に起きる（そこで鳴らすと誤爆になる）。
    ★ キーの分解は rsplit: シート名に "!" を含められる（Excel が禁じるのは : \\ / ? * [ ]）。
    """
    out: dict = {}
    for key, tup in (snap.get("cells") or {}).items():
        sheet, rc = key.rsplit("!", 1)
        r, c = rc.split(",")
        value = tup[0] if tup else None
        if value in (None, ""):
            continue
        out[(sheet, int(r), int(c))] = value
    return out


def _changed(before: dict, after: dict) -> list:
    """値が変わったセルの一覧 [((シート, 行, 列), 前, 後), ...]（表示順が安定するよう整列）。"""
    b, a = _values(before), _values(after)
    return [(k, b.get(k), a.get(k)) for k in sorted(set(b) | set(a)) if b.get(k) != a.get(k)]


def _samples(hits: list, *, cell_ref: Callable, fmt_value: Callable) -> str:
    """変更セルの実例（先頭 _MAX_SAMPLES 件）を「シート!B2: 前 → 後」の形で並べる。"""
    shown = hits[:_MAX_SAMPLES]
    body = "、".join(f"{sheet}!{cell_ref(r, c)}: {fmt_value(old)} → {fmt_value(new)}"
                     for (sheet, r, c), old, new in shown)
    more = f"、ほか{len(hits) - len(shown)}件" if len(hits) > len(shown) else ""
    return body + more


def _check_new_row_at_end(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable, **_kw):
    """前提: 書き込んだ行は before で空だった（＝末尾に**新しい**行を足したはず）。

    ★ 既存の合計行を潰した致命の検出そのもの。「表の下端をどこと判定したか」は問わない
    （それは別の根＝_scan_last_row_basic の領分）。ここが見るのは結果だけ ──
    値を書いた行に、適用前から何か入っていたか。
    """
    occupied = {(sheet, r) for (sheet, r, _c) in _values(before)}
    hits = [h for h in _changed(before, after) if (h[0][0], h[0][1]) in occupied]
    if not hits:
        return None
    return (f"★ 末尾に新しい行を足すはずが、既存の行の値を {len(hits)} 件書き換えました"
            f"（{_samples(hits, cell_ref=cell_ref, fmt_value=fmt_value)}）")


def _looks_like_own_prior_output(before: dict, sheet: str, expected_header) -> bool:
    """★★ 単位H: before のそのシートの 1 行目が、**その op 自身の出力の見出し**と一致するか。

    一致すれば「前回そこに書いたのは自分」＝作り直しであって、人の作ったものの破壊ではない。
    ★ 署名は実装（helpers/*.bas）から取る。想像で決めない ── 呼び出し側が渡す
    （このモジュールは ailine.py も .bas も知らない）。
    """
    if not expected_header:
        return False
    cells = before.get("cells") or {}
    for i, want in enumerate(expected_header):
        got = cells.get(f"{sheet}!1,{i + 1}")
        if (got[0] if got is not None else None) != want:
            return False
    return True


def _check_new_sheet(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable,
                     own_output_headers=None, **_kw):
    """前提: その名前のシートは before に存在しない（＝**新しい**シートを作ったはず）。

    ★ 既に在ったシートの値を書き換えた（消して作り直したものを含む）ら前提が破れている。
    新規シートに書いた分は before に無いシートなので、ここでは一切拾わない。

    ★★ 単位H（誤爆抑制の復元）: ただし、そのシートが **自分の前回の出力**だったなら
    前提は破れていない ── 2 回目の集計は正常な作り直しであって破壊ではない。
    実測（2026-08-19）: 単位F/G だけの状態では、元データが増えた後の 2 回目の AGGREGATE が
    exit 7 で止まっていた。判定材料は呼び出し側から来る見出し署名だけで、ここは形を知らない。
    """
    existing = set(before.get("sheets") or [])
    own = {s for s in existing
           if _looks_like_own_prior_output(before, s, (own_output_headers or {}).get(s))}
    hits = [h for h in _changed(before, after) if h[0][0] in existing and h[0][0] not in own]
    if not hits:
        return None
    sheets = sorted({h[0][0] for h in hits})
    names = "、".join(f"『{s}』" for s in sheets)
    return (f"★ 新しいシートを作るはずが、既存のシート{names}の値を {len(hits)} 件"
            f"書き換えました（{_samples(hits, cell_ref=cell_ref, fmt_value=fmt_value)}）")


def _check_format_only(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable, **_kw):
    """前提: 値が1つも変わらない（書式・罫線・列幅・埋め込みグラフだけを触ったはず）。"""
    hits = _changed(before, after)
    if not hits:
        return None
    return (f"★ 書式だけのはずが、セルの値が {len(hits)} 件変わりました"
            f"（{_samples(hits, cell_ref=cell_ref, fmt_value=fmt_value)}）")


def _check_value_multiset(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable, **_kw):
    """前提: 値の多重集合が保存される（行を挿入してずらす/並べ替えるだけ）。

    ★ 数式は比較から外す: 行が動けば参照も書き換わる（=D2-C2 → =D3-C3）ので、文字列として
    比べると必ず食い違う ── それは移動の正常な副作用であって破壊ではない。
    ★ 見るのは「消えた値」だけ。増えた値は破壊ではない（それは幽霊データ検出の領分）。
    """
    def bag(snap):
        return Counter(str(v) for v in _values(snap).values()
                        if not (isinstance(v, str) and v.startswith("=")))

    lost = bag(before) - bag(after)
    if not lost:
        return None
    total = sum(lost.values())
    shown = sorted(lost.elements())[:_MAX_SAMPLES]
    body = "、".join(fmt_value(v) for v in shown)
    more = f"、ほか{total - len(shown)}件" if total > len(shown) else ""
    return f"★ 行を動かすだけのはずが、値が {total} 件消えました（{body}{more}）"


# 書き込み領域の種類 → その前提を確かめる関数。★ キーの文字列は ailine.py の WRITE_* 定数
# （宣言表 OP_WRITE_TARGET が使う語彙）と一致していなければならない ── 一致は番人テスト
# （tests/test_write_precondition_unit.py）が検査する。ここに載っていない種類は
# NO_PRECONDITION に載せる＝「前提は無いと確認した」という宣言（忘れたのではない）。
PRECONDITIONS = {
    "new_row_at_end": _check_new_row_at_end,
    "new_sheet": _check_new_sheet,
    "format_only": _check_format_only,
    "row_shift": _check_value_multiset,
    "reorder": _check_value_multiset,
}

# 前提を持たない種類（既存の破壊の関所＝書き込み先列の既存値検知が守る側）。
NO_PRECONDITION = frozenset({"existing_column", "new_column"})


def check_write_preconditions_detail(writes, before: dict, after: dict, *,
                                     cell_ref: Callable, fmt_value: Callable,
                                     own_output_headers=None):
    """破れた前提を **(種類, 文言)** で返す（破れていなければ None）。

    ★ 単位G が種類を要る理由: 「前提が破れた」だけでは、どの宣言が嘘をついたのか分からない。
    中立化（「（既存シート『集計』の更新は意図どおりです）」）を黙らせてよいのは
    **new_sheet の前提が破れたとき**であって、format_only や reorder が破れたときではない。
    種類を返さずに「何か破れた」で中立化を止めると、関係の無い理由で肯定文を消すことになる。

    ★ 検査の順は writes の宣言順に従う（最初に破れたものを返す）。1 op が複数の前提を
    破ったときに 2 行出すことはしない ── 関所に渡すのは 1 行という約束を変えないため。
    """
    for kind in writes or ():
        check = PRECONDITIONS.get(kind)
        if check is None:
            continue
        message = check(before, after, cell_ref=cell_ref, fmt_value=fmt_value,
                        own_output_headers=own_output_headers)
        if message:
            return kind, message
    return None


def check_write_preconditions(writes, before: dict, after: dict, *,
                               cell_ref: Callable, fmt_value: Callable,
                               own_output_headers=None) -> str | None:
    """宣言した writes の前提が適用後の実測と食い違っていれば、その1行を返す（無ければ None）。

    writes: op の宣言（OP_WRITE_TARGET[op].writes）。呼び出し側が渡す ── ここは
            ailine.py を import しない（別プロジェクトへそのまま持ち出せる形を保つ）。
    before/after: ailine.snapshot() の dict（before は原本、after は適用済みのコピー）。
    cell_ref/fmt_value: 表示用の関数（ailine.py の _cell_ref / _fmt_cell_value）。同じ表記を
            2箇所に書かないため、呼び出し時点で受け取る（formula_health.py と同じ作法）。

    ★ 単位G 以降、本体は check_write_preconditions_detail（種類つき）。ここはその文言だけを
    返す薄い皮で、判定は 1 箇所しか無い（同じ検査を 2 箇所に書かない）。
    """
    detail = check_write_preconditions_detail(writes, before, after,
                                              cell_ref=cell_ref, fmt_value=fmt_value,
                                              own_output_headers=own_output_headers)
    return detail[1] if detail else None
