# -*- coding: utf-8 -*-
"""新しく作ったシートが**空**なら、頼まれたことは起きていない。

★★ なぜ在るか（2026-09-08・盲検 B が実害のある false ✓ として拾った）:

    依頼   「合計より上の行だけ抽出して新しいシートにコピーして」
    実行   操作:抽出 対象列:商品 条件:等しい 値:合計   ← 条件を誤変換
    事後条件 「4行中**0行**が一致 → 0行を抽出（値・型とも保存）」  ← ★ 自分で数えている
    出力   **✓ 機械検証済み**／ 実物は見出しだけの空シート

  ★ 検出は在って**帰結が無い**（同日の #13「壊れた式を原本へ書く」と同じ形）。
    抽出としては「0 行を正しく抽出した」ので事後条件は嘘をついていない ── だが
    依頼は「行を抜き出す」で、空の成果物は依頼が満たされていない。

★★ 「0 件なら落とす」ではない ── 掃き出して分かった（4 op を実際に呼んだ）:

    EXTRACT   0 行抽出  → pass  ★ 欠陥（成果物が空）
    DEDUP     重複 0 件 → pass  ← **正しい**（除くものが無いのは正常・表は残る）
    SET_WHERE 0 行一致  → fail  ← 既に守っている
    AGGREGATE 0 群      → fail  ← 既に守っている

  ★ 分かれ目は件数でなく**成果物**: 新しく作ったシートに見出し以外が 1 行も無いか。
    DEDUP の出力は 2 行あるので鳴らない。op 名を列挙しないので、
    新しい op を足しても自動で守られる。

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（他の関所と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations

from pathlib import Path

from ailine_core.book_view import BookView


def _has_data_rows(bv, sheet: str) -> bool:
    """見出し以外に、値の入った行が 1 行でもあるか。

    ★ 「見出し行」を機械が決められない場面でも効くよう、**2 行目以降に非空セルが
      1 つでもあるか**だけを見る（見出し行の推定に依存しない ── 依存すると、
      推定が外れた回に黙る）。"""
    ws = bv.sheet(sheet)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            v = cell.value
            if v is not None and str(v).strip() != "":
                return True
    return False


def empty_new_sheets(before_path, after_path) -> list:
    """適用で**新しく出来た**シートのうち、中身が見出しだけ（またはそれ以下）のものを返す。

    ★ 読めない回は空を返す（測れないものを鳴らさない）。
    """
    try:
        with BookView(Path(before_path)) as bb:
            before = set(bb.sheetnames)
        with BookView(Path(after_path)) as ab:
            names = list(ab.sheetnames)
            return [s for s in names if s not in before and not _has_data_rows(ab, s)]
    except Exception:
        return []
