# -*- coding: utf-8 -*-
"""分けた冊（`ailine split` の出力）を**後から・ファイルだけから**独立に検算する（2026-09-12）。

## なぜ在るか

書き手（`cmd_split`）は既に、書いた冊を読み戻して帰属・行数・Σ を検算している。
★ しかしその分母は **書き手自身の意図（`plan`）** だ。同じ頭で作った分母で測ると、
規則の取り違えはすり抜ける（[[feedback_independent_verification]]）。
`ailine verify` にはこの段が無く、`{"unsupported"}` を正直に返すだけだった。

## 何を「独立」と呼ぶか（★ ここを外すと恒真になる）

**規則（誰にどの行を渡すか）を再現しない。** `split_people.plan_split` を呼ばない。
分母は**元の冊**から取り、測るのは保存則と忠実性:

    ① 元の非空行が、出力の出所（元ファイル・元行）か検分の名指しの **どちらかに 1 回だけ**在る
    ② 出力の各行の値が、元の行と **1 セルずつ**一致する（規則を知らなくても壊れは分かる）
    ③ 金額: 配った ＋ 空欄 ＋ 複数担当 ＝ 元の合計（★ 分けない行は除く ── 合計行が入るため）
    ④ 出力に**元に無い行**が無い（捏造の側）

★ 検分シートからは**名指し（何行目か）だけ**を読む。検分が書いた**数は読まない** ──
  書き手の主張を分母にしたら検算にならない。
★ 検分が名指しした行のうち、こちらの分母に無いもの（全部空の行など）は**咎めない**。
  製品の「データ行とは何か」の規則を再現しない、という線を守るため。
★ 出所列の名前と検分の冊名（`_検分`）は**出力の形式**そのものなので共有する（規則ではない）。

## 測った（2026-09-12・検体 5 冊）

    素の出力            破れ 0 件（5 冊すべて・金額の和も閉じた）
    行を 1 つ消す        取り逃し + 金額の和が閉じない
    別の冊にも複製する   二重配布 + 金額の和が閉じない
    値を書き換える       値が元と違う
    金額を 1 円ずらす    値が元と違う + 金額の和が閉じない
    元に無い行番号を足す  元に無い行を配った（捏造）

★ 最初の版は検分を読まず、配らなかった行（空欄・合計行）を「取り逃し」と 11 件鳴らした。
  **製品ではなく検算器が粗かった** ── 赤が出たら、まず 1 冊を手で解く。
"""
from __future__ import annotations

import re

from ailine_core import filetypes, primitives, split_people, xml_readback
from ailine_core import stack

#: 検分の「種類」── 配らなかった理由。★ 金額の和に入れるのは 空欄 と 複数担当 だけ
#:   （分けない行には合計行が入るので、足すと二重に数える）。
NOT_DISTRIBUTED = ("空欄", "複数担当", "分けない行")
IN_THE_SUM = ("空欄", "複数担当")
_ROW_NO = re.compile(r"(\d+)\s*行目")
#: 金額の差をこれ以上離れたら破れとする（円）。
AMOUNT_TOLERANCE = 0.5


def _grid_rows(path, *, sheet_name=None) -> tuple:
    """(見出し行, 見出しの名前, {行番号: {列: 値}}) を返す。

    ★ 見出し行はこの器官の**自前の規則**で決める（「2 つ以上埋まった最初の行」）──
      製品の `multifile.find_header_row` を呼ぶと、見出しの取り違えが一緒に動いて消える。
      取り違えた場合は「見出しが元と合いません」で**鳴る**（黙って進まない）。
    """
    data = xml_readback.read_grid(path, sheet_name)
    grid = data.get("grid") or {}
    if not grid:
        return None, [], {}
    rows = {}
    for (r, c), v in grid.items():
        if v is None or str(v).strip() == "":
            continue
        rows.setdefault(r, {})[c] = v
    head_row = next((r for r in sorted(rows) if len(rows[r]) >= 2), None)
    if head_row is None:
        return None, [], {}
    names = {c: str(v).strip() for c, v in rows[head_row].items()}
    body = {r: cells for r, cells in rows.items() if r > head_row}
    return head_row, names, body


def named_rows(report_path) -> dict:
    """検分が名指しした {行番号: 種類}。★ 名指しだけを読み、書かれた**数は読まない**。"""
    out = {}
    head_row, names, body = _grid_rows(report_path)
    if head_row is None:
        return out
    kind_col = next((c for c, n in names.items() if n == split_people.REPORT_HEADERS[0]), None)
    target_col = next((c for c, n in names.items() if n == split_people.REPORT_HEADERS[1]), None)
    if kind_col is None or target_col is None:
        return out
    for cells in body.values():
        kind = str(cells.get(kind_col) or "").strip()
        if kind not in NOT_DISTRIBUTED:
            continue
        m = _ROW_NO.search(str(cells.get(target_col) or ""))
        if m:
            out[int(m.group(1))] = kind
    return out


def _values_match(a, b) -> bool:
    """値が同じか。★ 数と文字の行き違いだけは許す（書式は検算の対象ではない）。"""
    if a == b:
        return True
    if primitives.is_number(a) and primitives.is_number(b):
        return abs(float(a) - float(b)) <= 1e-9
    return str(a).strip() == str(b).strip()


def _amount(v):
    return float(v) if primitives.is_number(v) else None


def output_books(out_dir) -> list:
    """検算の対象になる冊（検分の冊と一時ファイルを除く）。"""
    suffix = filetypes.OPENPYXL_READABLE_SUFFIX
    return sorted(p for p in out_dir.glob("*" + suffix)
                  if p.stem != split_people.REPORT_STEM and not p.name.startswith("~$"))


def verify_split_folder(out_dir, source, amount_header: str | None = None) -> dict:
    """分けた冊のフォルダを元の冊と突き合わせる。

    戻り値: {"breaks": [(名前, 名指し)], "facts": {...}, "mismatch": bool} または
            {"unsupported": 理由}（検算できない形 ── 空虚な合格を名乗らない）。
    """
    report = out_dir / (split_people.REPORT_STEM + filetypes.OPENPYXL_READABLE_SUFFIX)
    books = output_books(out_dir)
    if not books:
        return {"unsupported": f"分けた冊が見つかりません: {out_dir}"}
    marked = [b for b in books if xml_readback.read_core_properties(b)[0] == split_people.CREATOR_MARK]
    if not marked:
        return {"unsupported": "このフォルダの冊に `ailine split` の印がありません"
                               "（人のファイルを検算したことにしない）。"}
    if len(marked) != len(books):
        return {"unsupported": "印のある冊と無い冊が混ざっています"
                               f"（印あり {len(marked)} / 全部 {len(books)}）── "
                               "検算の分母が決まらないので断ります。"}

    breaks = []
    src_head, src_names, src_rows = _grid_rows(source)
    if src_head is None:
        return {"unsupported": f"元の冊から見出しが見つかりません: {source}"}
    src_col_of = {n: c for c, n in src_names.items()}
    amount_col = src_col_of.get(amount_header) if amount_header else None
    if amount_header and amount_col is None:
        return {"unsupported": f"元の冊に「{amount_header}」の列がありません"
                               "（`--amount` に元の冊の見出しを渡してください）。"}

    seen, parts_amount = {}, 0.0
    for book in books:
        head_row, names, body = _grid_rows(book)
        if head_row is None:
            breaks.append(("読めない冊", book.name))
            continue
        col_of = {n: c for c, n in names.items()}
        missing = [n for n in stack.PROVENANCE_HEADERS if n not in col_of]
        if missing:
            breaks.append(("出所の列が無い", f"{book.name} → {missing}"))
            continue
        unknown = [n for n in names.values()
                   if n not in stack.PROVENANCE_HEADERS and n not in src_col_of]
        if unknown:
            breaks.append(("★ 見出しが元と合いません", f"{book.name} → {unknown}"))
        i_file, i_row = (col_of[n] for n in stack.PROVENANCE_HEADERS)
        for r, cells in body.items():
            raw_row = cells.get(i_row)
            if not primitives.is_number(raw_row):
                breaks.append(("元行が数でない", f"{book.name} {r} 行目 → {raw_row!r}"))
                continue
            n = int(float(raw_row))
            name = str(cells.get(i_file) or "")
            if name and name.split("/")[-1].split("\\")[-1] != source.name:
                breaks.append(("元ファイル名が違う", f"{book.name} {r} 行目 → {name!r}"))
            if n not in src_rows:
                breaks.append(("★ 元に無い行を配った（捏造）", f"{book.name} → 元 {n} 行目"))
                continue
            seen.setdefault(n, []).append(book.name)
            for c, value in cells.items():
                header = names.get(c)
                if header in stack.PROVENANCE_HEADERS or header not in src_col_of:
                    continue
                want = src_rows[n].get(src_col_of[header])
                if not _values_match(value, want):
                    breaks.append(("★ 値が元と違う",
                                   f"{book.name} 元 {n} 行目「{header}」: "
                                   f"{value!r} ≠ {want!r}"))
            if amount_header and amount_header in col_of:
                got = _amount(cells.get(col_of[amount_header]))
                if got is not None:
                    parts_amount += got

    named = named_rows(report)
    for n in sorted(src_rows):
        where, kind = seen.get(n, []), named.get(n)
        if not where and not kind:
            breaks.append(("★ 取り逃し（配られず名指しもされていない）", f"元 {n} 行目"))
        elif where and kind:
            breaks.append(("★ 配ったのに「配らなかった」と名指し",
                           f"元 {n} 行目 → {kind} / {where}"))
        elif len(where) > 1:
            breaks.append(("★ 二重配布", f"元 {n} 行目 → {where}"))

    facts = {"元の非空行": len(src_rows), "配られた行": len(seen),
             "検分が名指しした行": len(named), "出力の冊": len(books)}
    if amount_col is not None:
        whole = held = 0.0
        for n, cells in src_rows.items():
            v = _amount(cells.get(amount_col))
            if v is None or named.get(n) == NOT_DISTRIBUTED[2]:
                continue
            whole += v
            if named.get(n) in IN_THE_SUM:
                held += v
        facts.update({"金額の見出し": amount_header, "Σ元（分けない行を除く）": whole,
                      "Σ配った": parts_amount, "Σ名指し": held})
        if abs(whole - (parts_amount + held)) > AMOUNT_TOLERANCE:
            breaks.append(("★ 金額の和が閉じない",
                           f"元 {whole} ≠ 配った {parts_amount} ＋ 名指し {held}"))
    else:
        facts["金額"] = "測っていません（`--amount <見出し>` を渡すと和も検算します）"
    return {"breaks": breaks, "facts": facts, "mismatch": bool(breaks)}
