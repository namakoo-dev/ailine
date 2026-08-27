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
  | new_column       | 同じ内容（見出し・値とも一致）の列が既にブックに無い |
  | new_row_at_end   | 書き込んだ行は before で空だった         |
  | new_sheet        | その名前のシートは before に存在しない   |
  | existing_column  | （前提なし＝上書き前提・既存の関所）     |
  | format_only      | 値が1つも変わらない                     |
  | row_shift/reorder| 値の多重集合が保存される                 |
  | remove           | 前提なし（値が減るのが正しい・事後条件が並びで証明する） |

★★ 単位J: new_column を NO_PRECONDITION から外し、本物の前提を与えた。
  盲検 operator 査定の実測: title_rows.xlsx に「売上から原価を引いた利益の列」を作らせ、
  同じ依頼をもう一度実行 → F 列と同じ見出し・同じ値の列が G 列にもう一つでき、警告ゼロで
  ✓ 機械検証済み まで出た。「反映されたか不安でもう一回実行」は事務職の最もありがちな操作。
  真因: 既存の関所（_maybe_warn_target_overwrite）が守るのは「書き込み先の列に既存値が
  あるか」で、新規列は定義上ずっと空だから何も鳴らない。「同じ結果が既にブックに在るか」は
  誰も聞いていなかった。判定・誤検知回避の設計は _check_new_column の docstring 参照。

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

    ★ 限界（単位H の敵対検証で確認済み・2026-08-20）: 署名は**出所（provenance）ではなく
    形**でしか無い。見出しセル（例: 分類列名『部門』・『合計 - 金額』）がたまたま今回の
    リクエストと一致する人間の手作りシート（データ行は無関係）を、この判定は区別できない
    ―― データ行も真の作成者も一切見ないため。さらに ailine が書いた出力を人が手で編集
    した後に同じ集計を依頼すると、その手編集は見出しの一致だけで「前回の自分の出力」と
    誤認され、関所を経由せず黙って失われる。真の出所追跡（provenance tracking、例:
    書き込み時刻や作成者を刻む仕掛け）は未実装 ── 見出し一致は「たぶん自分」の弱い代理
    指標にすぎない。
    """
    if not expected_header:
        return False
    cells = before.get("cells") or {}
    for i, want in enumerate(expected_header):
        got = cells.get(f"{sheet}!1,{i + 1}")
        if (got[0] if got is not None else None) != want:
            return False
    return True


def own_prior_output_notice_lines(before: dict, after: dict, own_output_headers=None) -> list:
    """★★ 単位H 開示: `_looks_like_own_prior_output` が真になって関所（前提破れ）をスキップ
       した既存シートについて、**その理由をユーザーに開示する**1行ずつのリスト
       （無ければ空リスト）。

    ★ これは助言（advisory）ではない ── 関所が黙る理由の開示そのもの。単位H 導入前は
    「前提は破れていない」と判定した経緯が画面に一切出ず、完全な無言のまま exit 0 に
    進んでいた（上の docstring の限界も参照）。せめて「なぜ黙ったか」を1行で見せる。
    ★ 対象は「実際に変化があった」シートだけ（own と判定されても before/after で
    何も変わっていなければ、そもそも黙る理由を説明する必要が無い）。
    ★ 置き場所は呼び出し側 ── 前提検査の直後（単発・複合計画の両経路、
    ailine.py の `_maybe_own_prior_output_notice` 参照）。
    """
    if not own_output_headers:
        return []
    existing = set(before.get("sheets") or [])
    own = {s for s in existing if _looks_like_own_prior_output(before, s, own_output_headers.get(s))}
    if not own:
        return []
    changed_sheets = {h[0][0] for h in _changed(before, after) if h[0][0] in existing}
    return [f"（前回の出力『{s}』を作り直します）" for s in sorted(changed_sheets & own)]


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


def _columns_by_sheet(values: dict) -> dict:
    """_values() の出力 {(シート,行,列): 値} を、列単位 {(シート,列): {行: 値}} へ束ね直す。"""
    out: dict = {}
    for (sheet, row, col), v in values.items():
        out.setdefault((sheet, col), {})[row] = v
    return out


def _check_new_column(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable, **_kw):
    """前提: 同じ内容（見出しセルと全データ値の両方）の列が、そのシートに before の時点で
    既に無い（＝これは本当に**新しい**列であって、前回と同じ結果の作り直しではない）。

    ★★ 単位J: 実測（盲検 operator 査定）── 「売上-原価」の列を作る依頼をもう一度実行すると、
    見出しも値も完全に同じ列がもう1本でき、警告ゼロで ✓ が出た。既存の関所は「書き込み先に
    既存値があるか」しか見ておらず、新規列は定義上ずっと空だから何も鳴らなかった。
    「同じ結果が既にブックに在るか」を初めて聞くのがこの関数。

    ★ 見出しと値の**両方**一致を要求する（片方だけでは鳴らさない）:
      - 値だけ一致 → LOOKUP_FILL 等の転記系が参照元と同じ値の列を作るのは正当な動作。
        値だけで鳴らすと誤爆する。
      - 見出しだけ一致 → 中身が違えば正当な作り直し（同名でも計算し直した結果）。
        見出しだけで鳴らすと誤爆する。
      観測された事故は「同じ依頼を2回実行」で見出しも値も完全に同一 ── まずそこを確実に
      捕まえる（W10e/単位H と同じ「全部一致した時だけ鳴らす」誤検知回避の作法）。

    ★ 比較は snapshot の範囲内（MAX_ROWS で切れていれば見えている範囲）でよい。保守側に倒れる
    ── 見える範囲で一致したなら、見えない範囲まで一致を疑って鳴らさないよりは安全（W10c と同じ
    「切り詰めは前提を緩めない」判断）。
    """
    before_vals, after_vals = _values(before), _values(after)
    before_cols = _columns_by_sheet(before_vals)
    after_cols = _columns_by_sheet(after_vals)
    before_keys = {(s, c) for (s, _r, c) in before_vals}
    after_keys = {(s, c) for (s, _r, c) in after_vals}
    for sheet, col in sorted(after_keys - before_keys):
        new_data = after_cols[(sheet, col)]
        for bsheet, bcol in sorted(before_cols):
            if bsheet != sheet or bcol == col or before_cols[(bsheet, bcol)] != new_data:
                continue
            header = new_data.get(1)
            header_disp = str(header) if header is not None else "(見出し無し)"
            letter = cell_ref(1, bcol).rstrip("0123456789")
            return (f"★ 新しい列を作るはずが、既存の列『{header_disp}』({letter}) と"
                    "見出しも値も同一の列を作りました（同じ依頼を 2 回実行した可能性）")
    return None


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


def _check_single_cell(before: dict, after: dict, *, cell_ref: Callable, fmt_value: Callable, **_kw):
    """前提: **値が変わったセルはちょうど 1 つ**（宣言した 1 セルだけを書く op）。

    ★ 2026-08-27: この種類が生まれた理由は、1 セル書きで「空欄への一括書き込み」の助言が
      誤爆したこと。助言を黙らせるなら、**黙らせた分の保証をここで取り返す**
      ── 鳴らなくした代わりに、宣言（1 セル）と実体（変わった数）を突き合わせる。
    ★ 事後条件 check_set_cell_value も同じことを証明するが、あちらは**適用後のファイルを
      読み直して**判定する。こちらは**差分の写真**から判定する ── 出どころが違う
      2 つで見るのが要点で、片方が黙った時にもう片方が鳴る。
    """
    changed = [k for k in set(_values(before)) | set(_values(after))
                if _values(before).get(k) != _values(after).get(k)]
    if len(changed) <= 1:
        return None
    shown = "、".join(f"{s}!{cell_ref(r, c)}"
                      for s, r, c in sorted(changed)[:_MAX_SAMPLES])
    more = f"、ほか{len(changed) - _MAX_SAMPLES}件" if len(changed) > _MAX_SAMPLES else ""
    return (f"★ 1 セルだけ書くはずが、値の変わったセルが {len(changed)} 個あります"
            f"（{shown}{more}）")


# 書き込み領域の種類 → その前提を確かめる関数。★ キーの文字列は ailine.py の WRITE_* 定数
# （宣言表 OP_WRITE_TARGET が使う語彙）と一致していなければならない ── 一致は番人テスト
# （tests/test_write_precondition_unit.py）が検査する。ここに載っていない種類は
# NO_PRECONDITION に載せる＝「前提は無いと確認した」という宣言（忘れたのではない）。
PRECONDITIONS = {
    "new_column": _check_new_column,
    "new_row_at_end": _check_new_row_at_end,
    "new_sheet": _check_new_sheet,
    "format_only": _check_format_only,
    "row_shift": _check_value_multiset,
    "reorder": _check_value_multiset,
    "single_cell": _check_single_cell,
}

# 前提を持たない種類（既存の破壊の関所＝書き込み先列の既存値検知が守る側）。
# ★ 2026-08-26: "remove"（行/列の削除）は**値が減るのが正しい**ので、値の保存を前提に
#   できない。忘れたのではなく「前提は無いと確認した」側に置く ── 代わりに
#   check_delete_rows / check_delete_column が「残りが順序ごと一致すること」を証明し、
#   消した中身は必ず画面に出す（note_deleted）。
NO_PRECONDITION = frozenset({"existing_column", "remove"})


# ★ 位置で比べる前提（列を**番号で**突き合わせるもの）。位置がずれる回は使えない。
POSITION_BASED = frozenset({"new_column"})


def check_write_preconditions_detail(writes, before: dict, after: dict, *,
                                     cell_ref: Callable, fmt_value: Callable,
                                     own_output_headers=None,
                                     positions_shifted: bool = False):
    """破れた前提を **(種類, 文言)** で返す（破れていなければ None）。

    ★ 単位G が種類を要る理由: 「前提が破れた」だけでは、どの宣言が嘘をついたのか分からない。
    中立化（「（既存シート『集計』の更新は意図どおりです）」）を黙らせてよいのは
    **new_sheet の前提が破れたとき**であって、format_only や reorder が破れたときではない。
    種類を返さずに「何か破れた」で中立化を止めると、関係の無い理由で肯定文を消すことになる。

    ★ 検査の順は writes の宣言順に従う（最初に破れたものを返す）。1 op が複数の前提を
    破ったときに 2 行出すことはしない ── 関所に渡すのは 1 行という約束を変えないため。
    """
    # ★★ 2026-08-27（Namakoo が実測）: **位置がずれる op では、位置で比べる前提は使えない。**
    #   列を途中に挿すと右の列が 1 つずつずれるので、new_column の前提（「同じ内容の列が
    #   既に在るか」を位置で見る）が「ずれた既存列」を毎回『同じ列を 2 回作った』と誤報する。
    #   ★ 前提そのものが壊れているので、宣言（row_shift と new_column の同居）で外す ──
    #     op 名の if ではなく、**なぜ使えないか**が読める形で。
    #   ★ 外した分の保証: ADD_COLUMN は verify_dsl_args が「同名の列が既に在る」を先に断り、
    #     check_add_column が「挿した位置・空であること・他の列の不変」を証明する。
    kinds = list(writes or ())
    #   ★ 2026-08-27（2 度目・実測）: 位置がずれるのは**宣言**からだけでなく、
    #     その回の引数からも起きる（依頼文の位置指定で新しい列を動かした回）。
    #     宣言（row_shift + new_column）と実測（positions_shifted）の**どちらでも**外す。
    if positions_shifted or ("row_shift" in kinds and "new_column" in kinds):
        kinds = [k for k in kinds if k not in POSITION_BASED]
    for kind in kinds:
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
