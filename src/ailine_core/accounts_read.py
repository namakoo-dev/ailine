# -*- coding: utf-8 -*-
"""accounts_read — 仕訳の書き出し（CSV / xlsx）を「行の並び」にする読み手（需要③・2026-09-13）。
   設計: docs/DESIGN-20260913-経費の勘定科目を先例から引く.md §6.2

★ なぜ `accounts_core` と分けるか: あちらは**純関数・I/O なし**（同じ入力なら同じ計画）。
  ファイルを開くのはこちらだけ ── 書き手（`ailine.cmd_accounts`）と検算（`verify_accounts`）の
  **両方**がこの 1 つを呼ぶ（読み手を 2 つ書くと、片方だけ直る）。

★★ 文字コードは `csv_quarantine.detect_encoding` を**再利用する**（新しい判定器を書かない）。
  BOM → UTF-8 → cp932 の順で、EUC/ISO 系は名指しで断る。曖昧（UTF-8 でも cp932 でも
  復号できる）な冊は**必ず画面に出す**ため、`ambiguous` を結果に運ぶ。

★ 日付は読まない（原文の文字列をそのまま運ぶ）。数値も読み替えない ── CSV の値は
  文字のまま渡す（`match.normalize_key` の線で、数値 123 と文字 "123" は別の鍵になる）。
"""
from __future__ import annotations

import csv
from dataclasses import dataclass, field

import openpyxl

from ailine_core import accounts_core, csv_quarantine, filetypes

#: 1 冊から読む行数の上限（★ 仕訳帳は長い ── 打ち切ったら打ち切ったと言う）。
MAX_ROWS = 200000


@dataclass(frozen=True)
class JournalBook:
    """1 冊の読み結果（★ 読めなかったら `refused` に理由 ── 黙って空にしない）。

    rows:       [(行番号, [値, ...]), ...] ★ 見出し行を**含まない**データ行だけ
    header_row: 見出しの行番号（弥生＝見出し行が無い冊は None）
    """
    name: str
    path: str = ""
    header_row: int | None = None
    headers: list = field(default_factory=list)
    header_map: dict = field(default_factory=dict)
    rows: list = field(default_factory=list)
    width: int = 0
    encoding: str | None = None
    ambiguous: bool = False
    truncated: bool = False
    refused: str | None = None


def _raw_rows_from_csv(path) -> tuple:
    """CSV を (行の並び, 文字コード, 曖昧か) で返す。★ 判定器は再利用・自作しない。"""
    raw = path.read_bytes()
    result = csv_quarantine.detect_encoding(raw)
    text = raw.decode(result.encoding)
    rows = []
    for i, values in enumerate(csv.reader(text.splitlines()), start=1):
        if i > MAX_ROWS:
            return rows, result.encoding, result.ambiguous, True
        rows.append((i, list(values)))
    return rows, result.encoding, result.ambiguous, False


def _raw_rows_from_book(path) -> tuple:
    """xlsx を (行の並び, 打ち切ったか) で返す（1 枚目のシート・値だけ）。"""
    book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = book.worksheets[0]
        rows, truncated = [], False
        for i, values in enumerate(sheet.iter_rows(values_only=True), start=1):
            if i > MAX_ROWS:
                truncated = True
                break
            rows.append((i, list(values)))
        return rows, truncated
    finally:
        book.close()


def read_journal(path) -> JournalBook:
    """1 冊の仕訳（CSV / xlsx）を読み、列を解決して `JournalBook` を返す。

    ★ 断るのは 3 つ: 扱えない形式／文字コードが決められない／列が決まらない。
      どれも**見たもの**を名指しする（推測で先へ進まない）。
    """
    suffix = path.suffix.lower()
    encoding, ambiguous, truncated = None, False, False
    try:
        if suffix == filetypes.CSV_SUFFIX:
            raw_rows, encoding, ambiguous, truncated = _raw_rows_from_csv(path)
        elif suffix == filetypes.OPENPYXL_READABLE_SUFFIX:
            raw_rows, truncated = _raw_rows_from_book(path)
        else:
            return JournalBook(name=path.name, path=str(path),
                               refused=f"扱えない形式です（{path.name}）── 仕訳の書き出しは "
                                       f"{filetypes.CSV_SUFFIX} か "
                                       f"{filetypes.OPENPYXL_READABLE_SUFFIX} で渡してください")
    except csv_quarantine.UndecidableEncodingError as e:
        return JournalBook(name=path.name, path=str(path),
                           refused=f"文字コードが決められません（{path.name}）: {e}")
    except Exception as e:   # noqa: BLE001 ── 名指しして断る（1 冊で止めない側は呼び出し元）
        return JournalBook(name=path.name, path=str(path),
                           refused=f"読み込みに失敗しました（{type(e).__name__}）: {path.name}")

    header_row, headers, header_map, refusal = \
        accounts_core.resolve_accounts_columns(raw_rows)
    width = max((len(v) for _r, v in raw_rows), default=0)
    if refusal:
        return JournalBook(name=path.name, path=str(path), header_row=header_row,
                           headers=list(headers), width=width, encoding=encoding,
                           ambiguous=ambiguous, truncated=truncated,
                           refused=f"{path.name}: {refusal}")
    rows = [(r, v) for r, v in sorted(raw_rows, key=lambda rv: rv[0])
            if header_row is None or r > header_row]
    return JournalBook(name=path.name, path=str(path), header_row=header_row,
                       headers=list(headers), header_map=header_map, rows=rows,
                       width=width, encoding=encoding, ambiguous=ambiguous,
                       truncated=truncated)


def past_pool(books) -> dict:
    """読めた過去の冊を `plan_accounts` の `past_rows_by_file` の形へ畳む。

    ★ ファイルごとに列を解決した結果をそのまま運ぶ（同じソフトでも列順が同じとは限らない）。
    """
    return {book.name: {"header_map": dict(book.header_map), "rows": list(book.rows)}
            for book in books if not book.refused}
