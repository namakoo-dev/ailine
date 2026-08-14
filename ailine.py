#!/usr/bin/env python
"""ailine — 自然言語のタスクを、ローカル LLM が LibreOffice Basic に書き起こし、
   basrun で文書に適用し、★ 効果を読み戻して検証する（「走った ≠ できた」）。

    ailine run  <book> "<タスク>"          生成 → 適用(コピー) → 変化検証 → 修復 → 差分表示
    ailine run  <book> "<タスク>" --dry     生成して見せるだけ（適用しない・レビュー用）
    ailine stop                             起動した LibreOffice を落とす（basrun に委譲）

## 設計判断（basrun_spike 2026-08-10 の実証に基づく）

- **モデル非依存**: `--model` で差し替え（既定 qwen2.5-coder:7b）。天井はモデルの大きさでなく
  「参照例の供給＋効果の検証」で上げる（7B が正解例1本で苦手層 0%→67% になった）。
- ★ **検証をループに**: 適用の前後で文書が変化したかを見る **no-op ガード**。
  LibreOffice + LLM は「実行時エラー無しで成功と報告し、実際は何もしない」ことがある
  （もっともらしい UNO の幻覚）。変化ゼロなら失敗として修復に回す。
- ★ **正規化パス（2026-08-14 修正・製品の心臓）**: before スナップショットの前に、
  コピーを LibreOffice で一度（空マクロで）開いて保存する。openpyxl 製ブックは LO の
  初回保存で行高（時に列幅）を実体化する副作用があり、これを先に済ませておかないと
  no-op ガードが偽陽性になる（何もしないマクロでも「変化した」と誤って成功表示していた）。
  コストは LO 往復 1 回（数秒）— 正しさ優先で受け入れる。
- **コピー安全**: 原本は触らず `<book>.out.xlsx` に適用する（`--inplace` で上書き）。壊さない。
- **参照ライブラリ**: `refs/*.bas` を few-shot に供給。苦手層（新シート・色）を補う。
  並べ替え・グラフ・ピボットなどの難所は `helpers/*.bas` の検証済みヘルパを `Call` で呼ばせる。
- **レビュー導線**: 生成した .bas と、変わったセルの差分を必ず表示する。

## 正直な限界

- LibreOffice Basic + UNO は学習データが薄く、珍しい操作は外しやすい。参照とヘルパで補う設計
  （太字も当初「環境不可」と誤断したが、実際は `CharWeight`+`CharWeightAsian` の native 書きで解決済み）。
- ローカル LLM(ollama) と LibreOffice(basrun 経由) が要る。**外部送信はしない。**
- no-op ガードが保証するのは「変化したこと」だけ。「**正しいか**」は差分を人が見て判断する。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
import zipfile
from datetime import datetime, timezone
from pathlib import Path

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    sys.exit("openpyxl が要る:  pip install openpyxl")

HERE = Path(__file__).resolve().parent
DEFAULT_REFS = HERE / "refs"
DEFAULT_HELPERS = HERE / "helpers"
OLLAMA = os.environ.get("OLLAMA_HOST", "http://localhost:11434").rstrip("/")
DEFAULT_MODEL = os.environ.get("AILINE_MODEL", "qwen2.5-coder:7b")
DEFAULT_APPLY_TIMEOUT = 180.0  # M1: 暴走マクロで無限ハングしないよう既定 ON（--timeout 0 で無効化）

HISTORY_DIR = Path.home() / ".ailine"
HISTORY_FILE = HISTORY_DIR / "history.jsonl"
BACKUP_DIR = HISTORY_DIR / "backups"
DEFAULT_KEEP_BACKUPS = 10   # M2c: book ごとにこの世代数を超えたら古い順に削除する


def _find_basrun_path() -> Path | None:
    """basrun.py の場所。環境変数 BASRUN > ailine と並びの checkout の順で探す。
       見つからなければ None（sys.exit しない版。doctor から非致命的に使う）。"""
    env = os.environ.get("BASRUN")
    if env:
        p = Path(env)
        return p if p.exists() else None
    for name in ("basrun", "nagi-bas"):  # 公開 repo 名 / 作者ローカルの旧ディレクトリ名
        p = HERE.parent / name / "basrun.py"
        if p.exists():
            return p
    return None


def basrun_path() -> Path:
    """basrun.py の場所。無ければ理由つきで落とす（run から使う致命版）。"""
    p = _find_basrun_path()
    if p is None:
        sys.exit("basrun.py が見つからない: ailine と並びに"
                 " https://github.com/namakoo-dev/basrun を clone するか、"
                 "環境変数 BASRUN でパスを指定する")
    return p


# ---------------------------------------------------------------------------
# 進捗表示（M1: 生成中の完全沈黙を解消。凝った演出はしない — Windows コンソール安全第一）
# ---------------------------------------------------------------------------

def _fmt_elapsed(seconds: float) -> str:
    """経過秒を表示用に整形する（例: 12.34 → '(12.3s)'）。"""
    return f"({seconds:.1f}s)"


def progress_start(label: str) -> float:
    """進捗の開始行を stderr に出す（改行しない＝完了時に経過秒を追記する）。
       --json のときも常に stderr（stdout の機械出力を汚さない）。
       開始時刻(time.monotonic())を返す。"""
    print(label, end="", file=sys.stderr, flush=True)
    return time.monotonic()


def progress_end(start: float) -> None:
    """開始行に経過秒を追記して改行する。"""
    print(" " + _fmt_elapsed(time.monotonic() - start), file=sys.stderr)

CONTRACT = """あなたは LibreOffice Basic を書く。出力は .bas のコードだけ。説明・markdown 柵は禁止。

厳守する契約:
- 先頭に `Option VBASupport 1` と `Option Explicit`。
- 手続きは **ちょうど1つ**、必ず `Sub Run(oDoc As Object)` という名前と署名。
- `ThisComponent` は使わない。文書は引数 oDoc で受け取る。
- シート: `oDoc.Sheets.getByIndex(0)`。セルは 0 起点 `getCellByPosition(列, 行)`。
  値は `.getValue()/.setValue()`、文字は `.getString()/.setString()`。
- 数値書式は `oDoc.getNumberFormats()` の queryKey/addNew を取り `範囲.NumberFormat = nFmt`。
- 列を文字("A")で指す API は使わない（例外で静かに止まる）。数値の列番号だけ。
"""

# セルの状態を snapshot する際の使用範囲の上限（病的に巨大な文書での暴走を防ぐ）
MAX_ROWS = 1000
MAX_COLS = 64


# ---------------------------------------------------------------------------
# ローカル LLM（ollama）
# ---------------------------------------------------------------------------

def ollama_generate(model: str, messages: list, temperature: float = 0.2) -> str:
    body = json.dumps({
        "model": model, "messages": messages, "stream": False,
        "options": {"temperature": temperature, "num_predict": 1600, "num_ctx": 8192},
    }).encode()
    req = urllib.request.Request(f"{OLLAMA}/api/chat", data=body,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            d = json.load(r)
    except urllib.error.HTTPError as e:
        # ★ HTTPError は URLError のサブクラスなので先に拾う。
        #   404 は「繋がっているがモデルが無い」で、接続不能とは原因も対処も別。
        if e.code == 404:
            sys.exit(f"ollama にモデル '{model}' が見つからない (HTTP 404)。\n"
                     f"★ `ollama pull {model}` で取得してから再実行して。")
        sys.exit(f"ollama がエラーを返した ({OLLAMA}): HTTP {e.code} {e.reason}")
    except urllib.error.URLError as e:
        sys.exit(f"ollama に繋がらない ({OLLAMA}): {e}\n"
                 "★ `ollama serve` が動いているか確認。外部送信はしない設計。")
    return d.get("message", {}).get("content", "")


def load_refs(refs_dir: Path) -> str:
    """refs/*.bas を few-shot テキストに連結する。"""
    if not refs_dir.is_dir():
        return ""
    chunks = []
    for f in sorted(refs_dir.glob("*.bas")):
        chunks.append(f"--- 参考例: {f.stem} ---\n{f.read_text(encoding='utf-8').strip()}")
    if not chunks:
        return ""
    return "\n\nこれらは正しい書き方の参考（別タスク）:\n" + "\n".join(chunks) + "\n--- 参考ここまで ---\n"


def load_helpers(helpers_dir: Path) -> tuple:
    """helpers/*.bas を (プロンプト用カタログ, ファイル一覧) にする。
       ★ 難所（ソートの ContainsHeader 等）は人が検証したヘルパに閉じ込め、
          モデルには『Call で呼ぶだけ』させる。arcane 層の確度を上げる要。"""
    files = sorted(helpers_dir.glob("*.bas")) if helpers_dir.is_dir() else []
    if not files:
        return "", []
    srcs = "\n".join(f.read_text(encoding="utf-8").strip() for f in files)
    catalog = (
        "\n\n## 定義済みヘルパ（★ 呼ぶだけ・再定義しない）\n"
        "arcane な操作（並べ替え等）は、自分で書かず次のヘルパを使うこと。\n"
        "★ 呼び方は必ず `Call 名前(引数)` の形（Call を付ける。括弧つきで Call 無しは誤動作する）。\n"
        "★ ヘルパの中身は絶対に書き写すな（SummaryTable 等が長くても）。必ず `Call 名前(...)` の1行だけで呼ぶ。\n"
        "例: 金額が列1なら、金額で降順に並べ替え → `Call SortByColumn(oDoc, 1, False)`\n"
        "例: 金額(列1)の棒グラフ（項目名は先頭列に自動）→ `Call InsertBarChart(oDoc, 1)`\n"
        "例: A1とB1を結合 → `Call MergeCells(oDoc, 0, 0, 1, 0)`\n"
        "例: 先頭データ行(2行目)の前に1行挿入 → `Call InsertRows(oDoc, 1, 1)`\n"
        "例: 表に罫線を引く → `Call DrawTableBorders(oDoc)`\n"
        "例: 各列の幅を内容に合わせる → `Call AutoFitColumns(oDoc)`\n"
        "例: C列(列2)に、商品名(列0)をキーに『単価表』から値を引く（VLOOKUP相当）"
        " → `Call VLookupFromTable(oDoc, 0, 2, \"単価表\")`（参照表は 列0=キー・列1=値）\n"
        "例: 『ピボット』で部門(列0)ごとに金額(列1)を集計（本物の DataPilot・Excel で操作可）"
        " → `Call PivotSum(oDoc, 0, 1)`\n"
        "例: 『集計表／まとめ』を作る＝部門(列0)ごとの金額(列1)を見栄えのする普通の表に"
        "（罫線・カンマ・太字つき。★『ピボット』と明示されない集計は基本こちら）"
        " → `Call SummaryTable(oDoc, 0, 1)`\n"
        "例: 見出し行(行0, 列0〜4)を太字に → `Call StyleBold(oDoc, 0, 0, 4, 0)`\n"
        f"--- 定義済み（この通り既に存在する。再定義するな）---\n{srcs}\n--- ここまで ---\n")
    return catalog, files


def extract_bas(text: str) -> str:
    m = re.search(r"```(?:\w+)?\s*(.*?)```", text, re.S)
    return (m.group(1) if m else text).strip()


def valid_signature(code: str) -> bool:
    return re.search(r"Sub\s+Run\s*\(\s*oDoc\s+As\s+Object\s*\)", code, re.I) is not None


#  ★ 開始の "Sub 名前(...)" だけを数える（"End Sub" 内の "Sub" を誤って開始として
#     二重カウントしないよう、識別子が続く形に絞る）。
_SUB_OPEN_RE = re.compile(r"\bSub\s+[A-Za-z_]\w*", re.I)
_ENDSUB_RE = re.compile(r"\bEnd\s+Sub\b", re.I)


def is_truncated_code(code: str) -> bool:
    """生成 Basic が途中で切断された兆候を、構造だけを見て検出する（M2a）。
       ★ 正しさは判定しない。CONTRACT は『手続きはちょうど1つ、必ず End Sub で閉じる』
       前提のため、Sub/End Sub の対応と『最後の行がちょうど End Sub か』だけを見れば、
       開き括弧や識別子で途中生成が止まった典型パターン（＝最後の行が End Sub でない）
       を漏れなく拾える。bad_signature（署名自体が無い）とは別の失敗分類。"""
    stripped = code.strip()
    if not stripped:
        return True
    sub_count = len(_SUB_OPEN_RE.findall(stripped))
    endsub_count = len(_ENDSUB_RE.findall(stripped))
    if sub_count == 0 or sub_count != endsub_count:
        return True
    last_line = stripped.splitlines()[-1].strip()
    return not bool(_ENDSUB_RE.fullmatch(last_line))


# ---------------------------------------------------------------------------
# 文書の説明と snapshot（検証の土台）
# ---------------------------------------------------------------------------

def describe_book(path: Path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True)
    lines = [f"シート一覧: {wb.sheetnames}（1枚目 = {wb.sheetnames[0]!r}）"]
    ws = wb[wb.sheetnames[0]]
    nrow, ncol = ws.max_row or 0, ws.max_column or 0
    lines.append(f"1枚目のデータ範囲: 約 {nrow} 行 x {ncol} 列（列は 0 起点で 0..{max(ncol-1,0)}）。")
    headers = []
    for c in range(1, min(ncol, MAX_COLS) + 1):
        v = ws.cell(row=1, column=1 + (c - 1)).value
        if v not in (None, ""):
            headers.append(f"列{c-1}={v}")
    if headers:
        lines.append("行0(見出し): " + ", ".join(headers))
    lines.append("行1以降がデータ。")
    wb.close()
    return "\n".join(lines)


def book_columns(path: Path) -> dict:
    """全シートの見出し行（1行目の値）を {シート名: [列名,...]} で返す。
       ★ describe_book は1枚目だけの人間可読版。こちらは M2b の翻訳・検証が使う
       機械可読の接地情報（列は最初の空欄で打ち切る＝連続した見出しだけを列とみなす）。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        ncol = min(ws.max_column or 0, MAX_COLS)
        headers = []
        for c in range(1, ncol + 1):
            v = ws.cell(row=1, column=c).value
            if v in (None, ""):
                break
            headers.append(str(v))
        out[name] = headers
    wb.close()
    return out


def build_book_meta(path: Path) -> dict:
    """{"sheets": [...], "headers": {シート名: [列名,...]}}。M2b 翻訳・検証の接地情報。"""
    headers = book_columns(path)
    return {"sheets": list(headers.keys()), "headers": headers}


def _charts_count(path: Path) -> int:
    try:
        with zipfile.ZipFile(path) as z:
            return sum(1 for n in z.namelist()
                       if "chart" in n.lower() and n.lower().endswith(".xml")
                       and "/charts/chart" in n.lower())
    except Exception:
        return 0


def snapshot(path: Path) -> dict:
    """文書の状態を取る。★ 値/数値書式/背景色/太字 に加え、罫線・結合・列幅・行高・
       水平配置も捉える。これで『書式のみ・罫線のみ・列幅のみ・結合のみ・中央揃えのみ』
       の変更も『変化した』と検出でき、no-op 誤検出（＝効いているのに失敗扱い）を防ぐ。"""
    wb = openpyxl.load_workbook(path)
    snap = {"sheets": list(wb.sheetnames), "charts": _charts_count(path),
            "cells": {}, "merges": {}, "colw": {}, "rowh": {}}
    for name in wb.sheetnames:
        ws = wb[name]
        nrow = min(ws.max_row or 0, MAX_ROWS)
        ncol = min(ws.max_column or 0, MAX_COLS)
        for r in range(1, nrow + 1):
            for c in range(1, ncol + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                fill = None
                if cell.fill is not None and cell.fill.patternType:
                    fill = str(cell.fill.start_color.rgb)
                bold = bool(cell.font.bold) if cell.font else False
                numfmt = cell.number_format
                bd = cell.border
                bsig = (bd.left.style, bd.right.style, bd.top.style, bd.bottom.style) if bd else None
                if bsig == (None, None, None, None):
                    bsig = None
                # ★ 水平配置（中央揃え等）。既定は None/'general' 扱い。
                align = cell.alignment.horizontal if cell.alignment else None
                if align == "general":
                    align = None
                if (val in (None, "") and fill is None and not bold
                        and numfmt == "General" and bsig is None and align is None):
                    continue
                snap["cells"][f"{name}!{r},{c}"] = (val, numfmt, fill, bold, bsig, align)
        snap["merges"][name] = sorted(str(rng) for rng in ws.merged_cells.ranges)
        snap["colw"][name] = {k: round(d.width, 2) for k, d in ws.column_dimensions.items() if d.width}
        snap["rowh"][name] = {k: round(d.height, 2) for k, d in ws.row_dimensions.items() if d.height}
    wb.close()
    return snap


# ★ snapshot() の cell tuple の並びと対応する日本語ラベル（M1: 生 tuple 漏れの解消）。
_CELL_FIELD_LABELS = ("値", "数値書式", "背景", "太字", "罫線", "配置")
# snapshot() は完全に既定状態のセルを cells 辞書に入れない（skip 条件参照）。
# 片側が辞書に無い(=完全既定)ときは、この既定 tuple で埋める（全 None で埋めると
# 数値書式が『None→General』という偽の差分を作ってしまう＝実測で見つけた罠）。
_CELL_DEFAULT = (None, "General", None, False, None, None)


def _cell_ref(row: int, col: int) -> str:
    """0起点の内部行/列表記(r,c)を、人が読める A1 形式（例: B2）にする。"""
    return f"{get_column_letter(col)}{row}"


def _fmt_cell_value(v) -> str:
    """人が読める値表示（None は '(空)'、文字列はクォート）。"""
    if v is None:
        return "(空)"
    if isinstance(v, str):
        return f"'{v}'"
    return str(v)


def describe_cell_change(before: tuple | None, after: tuple | None) -> str:
    """(値,数値書式,背景,太字,罫線,配置) の tuple 差分を、★ 変わったフィールドだけ
       日本語ラベルで列挙する人間可読の文字列にする（tuple の生表示は出さない）。
       例: 値 'りんご'→'リンゴ', 太字 (空)→True"""
    b = before if before is not None else _CELL_DEFAULT
    a = after if after is not None else _CELL_DEFAULT
    parts = []
    for label, bv, av in zip(_CELL_FIELD_LABELS, b, a):
        if bv != av:
            parts.append(f"{label} {_fmt_cell_value(bv)}→{_fmt_cell_value(av)}")
    return ", ".join(parts) if parts else "(差分なし)"


def diff_snapshots(before: dict, after: dict) -> tuple:
    """(changed: bool, lines: [str])。人が読める変更点も返す。
       セル値/書式/色/太字/罫線・結合・列幅・行高・シート・グラフの変化を見る。"""
    lines = []
    added = [s for s in after["sheets"] if s not in before["sheets"]]
    removed = [s for s in before["sheets"] if s not in after["sheets"]]
    if added:
        lines.append(f"＋シート追加: {added}")
    if removed:
        lines.append(f"－シート削除: {removed}")
    if after["charts"] != before["charts"]:
        lines.append(f"＊グラフ数: {before['charts']} → {after['charts']}")

    # 結合セル
    merge_changes = 0
    for name in after.get("merges", {}):
        b = set(before.get("merges", {}).get(name, []))
        a = set(after["merges"][name])
        for m in sorted(a - b):
            merge_changes += 1; lines.append(f"＋結合 {name}!{m}")
        for m in sorted(b - a):
            merge_changes += 1; lines.append(f"－結合解除 {name}!{m}")

    # 列幅・行高（変わったシート数で示す）
    dim_changes = 0
    for key, label in (("colw", "列幅"), ("rowh", "行高")):
        for name in after.get(key, {}):
            if before.get(key, {}).get(name) != after[key].get(name):
                dim_changes += 1; lines.append(f"＊{label}変更: {name}")

    # セル（値/書式/色/太字/罫線）－ シートごとに自前の見出し→明細（他の種別と揃える）
    keys = set(before["cells"]) | set(after["cells"])
    changed_by_sheet: dict = {}
    cell_changes = 0
    for k in sorted(keys):
        b = before["cells"].get(k)
        a = after["cells"].get(k)
        if b == a:
            continue
        cell_changes += 1
        sheet, rc = k.split("!", 1)
        r_str, c_str = rc.split(",")
        ref = _cell_ref(int(r_str), int(c_str))
        changed_by_sheet.setdefault(sheet, []).append(f"  {ref}: {describe_cell_change(b, a)}")

    shown = 0
    for sheet, clines in changed_by_sheet.items():
        if shown >= 12:  # 全部は出さない。多いときは件数で示す
            break
        lines.append(f"＊セル値変更: {sheet}")
        for cl in clines:
            if shown >= 12:
                break
            lines.append(cl)
            shown += 1
    if cell_changes > shown:
        lines.append(f"  …ほか {cell_changes - shown} セル")

    changed = bool(added or removed or (after["charts"] != before["charts"])
                   or merge_changes or dim_changes or cell_changes)
    return changed, lines


# ---------------------------------------------------------------------------
# 疑わしい変化の機械検出（M2a: 冷間再監査が実測した「✓ の下の失敗」への対抗）
#
# ★ 全部「差分の後の助言」であってブロックしない。誤検知を恐れて条件は保守的に
#   （両条件とも『変更セルの全部がそれに該当』した時だけ発火）。
# ---------------------------------------------------------------------------

def _changed_cells(before: dict, after: dict) -> list:
    """(sheet, row, col) のリスト。値/書式/色/太字/罫線/配置のいずれかが変わったセル。"""
    keys = set(before["cells"]) | set(after["cells"])
    out = []
    for k in keys:
        if before["cells"].get(k) != after["cells"].get(k):
            sheet, rc = k.split("!", 1)
            r_str, c_str = rc.split(",")
            out.append((sheet, int(r_str), int(c_str)))
    return out


def _value_changed_cells(before: dict, after: dict) -> list:
    """(sheet, row, col) のリスト。★ M2c: 値(idx0)が実際に変わったセルだけ
       （書式(罫線・中央揃え等)だけが変わったセルは含めない）。
       幽霊データ/一様埋め検出を『値変更』の部分集合だけで評価するための土台。
       ★ 冷間再監査3回目の実測: 罫線+中央揃え+0埋めが混在すると、書式だけ変わった
       セル（値は不変）が uniform 判定の対象に紛れ込み、一様埋めが見逃されていた。"""
    keys = set(before["cells"]) | set(after["cells"])
    out = []
    for k in keys:
        b = before["cells"].get(k)
        a = after["cells"].get(k)
        b_val = b[0] if b is not None else None
        a_val = a[0] if a is not None else None
        if b_val != a_val:
            sheet, rc = k.split("!", 1)
            r_str, c_str = rc.split(",")
            out.append((sheet, int(r_str), int(c_str)))
    return out


def _used_range(before: dict, sheet: str) -> tuple | None:
    """原本(before snapshot)においてそのシートで実際に値が入っていたセルの矩形
       (min_row, max_row, min_col, max_col)。★ 近似: 値(idx0)が非空のセルだけを対象に
       する（書式のみのセルは『データ』に数えない）。値が1つも無ければ None。"""
    rows, cols = [], []
    prefix = sheet + "!"
    for k, v in before["cells"].items():
        if not k.startswith(prefix):
            continue
        if v[0] in (None, ""):
            continue
        r_str, c_str = k[len(prefix):].split(",")
        rows.append(int(r_str)); cols.append(int(c_str))
    if not rows:
        return None
    return (min(rows), max(rows), min(cols), max(cols))


def detect_ghost_data(before: dict, after: dict) -> str | None:
    """★ 幽霊データ検出: 変更セルが全部、原本の使用範囲（データが存在した矩形）の
       外に集中している場合だけ疑わしい旨を返す。1セルでも範囲内なら何も言わない
       （保守的。使用範囲が不明なシートが混ざる場合も判定を保留する）。
       ★ M2c: 判定対象は『値変更』の部分集合だけ（書式のみの変更は無視・保守性は部分集合内で維持）。"""
    changed = _value_changed_cells(before, after)
    if not changed:
        return None
    outside = []
    for sheet, r, c in changed:
        rect = _used_range(before, sheet)
        if rect is None:
            return None  # このシートの原本データ範囲が不明 → 判定を保留
        min_r, max_r, min_c, max_c = rect
        if min_r <= r <= max_r and min_c <= c <= max_c:
            return None  # 1つでも範囲内 → 発火しない
        outside.append((r, c))
    rows = [r for r, _ in outside]
    cols = [c for _, c in outside]
    top_left = _cell_ref(min(rows), min(cols))
    bot_right = _cell_ref(max(rows), max(cols))
    span = top_left if len(outside) == 1 else f"{top_left}:{bot_right}"
    return f"★ 疑わしい: 変更が元データの範囲外です（{span}）"


def detect_uniform_fill(before: dict, after: dict) -> str | None:
    """★ 一様埋め検出: 変更セルの全部で『変化前が空欄』かつ『変化後が全部同一値』
       （特に 0/空文字）の場合だけ疑わしい旨を返す（保守的）。
       ★ M2c: 判定対象は『値変更』の部分集合だけ（罫線・中央揃えなど書式のみが変わった
       セルは対象外にする — 混ざっていると後方の値だけ均一でも見逃していた実測不具合の修正）。"""
    keys = set(before["cells"]) | set(after["cells"])
    after_vals = []
    for k in keys:
        b = before["cells"].get(k)
        a = after["cells"].get(k)
        b_val = b[0] if b is not None else None
        a_val = a[0] if a is not None else None
        if b_val == a_val:
            continue  # 値は変わっていない（書式だけの変更は対象外）
        if b_val not in (None, ""):
            return None  # 変化前が空欄でないセルが1つでもある → 発火しない
        after_vals.append(a_val)
    if not after_vals:
        return None
    if len(set(after_vals)) != 1:
        return None
    val = after_vals[0]
    if val in (None, ""):
        return None  # 空欄→空欄は『埋めた』ことにならない
    return f"★ 疑わしい: 空欄への同一値の一括書き込みです（値 {_fmt_cell_value(val)} × {len(after_vals)} セル）"


def _data_row_count(before: dict, sheet: str, key_col: int) -> int:
    """★ 近似: そのシートの原本使用範囲のうち見出し行(1行目)を除いた行で、
       key_col（左隣接のキー列）が埋まっている行数をデータ行数の目安とする。
       結合セルや空行の扱いまでは厳密化しない、素直な近似。"""
    rect = _used_range(before, sheet)
    if rect is None:
        return 0
    min_r, max_r, _min_c, _max_c = rect
    start = max(min_r + 1, 2)
    prefix = sheet + "!"
    count = 0
    for r in range(start, max_r + 1):
        v = before["cells"].get(f"{prefix}{r},{key_col}")
        if v is not None and v[0] not in (None, ""):
            count += 1
    return count


def count_reconciliation(before: dict, after: dict) -> str | None:
    """変更が単一シート・単一列に集中している場合だけ「データ N 行のうち M 行を変更」
       を添える（りんご欠落型のような『1行だけ抜けている』変更を1秒で見えるように）。"""
    changed = _changed_cells(before, after)
    if not changed:
        return None
    sheets = {s for s, _, _ in changed}
    cols = {c for _, _, c in changed}
    if len(sheets) != 1 or len(cols) != 1:
        return None
    sheet = next(iter(sheets))
    col = next(iter(cols))
    rect = _used_range(before, sheet)
    if rect is None:
        return None
    _min_r, _max_r, min_c, _max_c = rect
    key_col = col - 1 if col > min_c else col
    data_rows = _data_row_count(before, sheet, key_col)
    changed_rows = len({r for _, r, _ in changed})
    unchanged_rows = max(data_rows - changed_rows, 0)
    col_letter = get_column_letter(col)
    return (f"列 {col_letter}: データ {data_rows} 行のうち {changed_rows} 行を変更"
            f"（{unchanged_rows} 行は未変更）")


# --- 依頼文言と変更範囲の重なりチェック（保守的・明示言及がある時だけ）--------

_RE_COL_KANJI_NUM = re.compile(r"列\s*(\d+)")
_RE_COL_LETTER_PREFIX = re.compile(r"列\s*([A-Za-z]{1,3})(?![A-Za-z0-9])")
_RE_COL_LETTER_SUFFIX = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]{1,3})\s*列")
_RE_ROW = re.compile(r"行\s*(\d+)")


def extract_task_mentions(task: str, sheet_names: list) -> dict:
    """タスク文言から明示的な言及だけを正規表現で抜き出す（保守的・誤検知回避優先）。
       戻り値: {"cols": {1起点の列番号(文字表記=曖昧さなし)}, "digit_cols": {数字表記の生の値},
               "rows": {1起点の行番号}, "sheets": {原本に実在するシート名}}
       ★ 列の数字表記(『列2』)は曖昧 — 本ツールの規約(README・ヘルパ)は 0 起点だが、
       人が普段言うのは 1 起点。どちらか一方に断定すると正しい結果へ誤警報を出す
       (2026-08-14 再監査 2 回目で実測:『在庫(列2)』= 0 起点 C 列への正しい変更に
       B 列不在の★が出た)。生の値のまま保持し、照合側で両解釈を許す。"""
    cols = set()
    digit_cols = {int(m.group(1)) for m in _RE_COL_KANJI_NUM.finditer(task)
                  if int(m.group(1)) >= 0}
    for pat in (_RE_COL_LETTER_PREFIX, _RE_COL_LETTER_SUFFIX):
        for m in pat.finditer(task):
            try:
                cols.add(column_index_from_string(m.group(1).upper()))
            except ValueError:
                pass
    rows = {int(m.group(1)) for m in _RE_ROW.finditer(task) if int(m.group(1)) >= 1}
    sheets = {s for s in sheet_names if s and s in task}
    return {"cols": cols, "digit_cols": digit_cols, "rows": rows, "sheets": sheets}


def _changed_sheets(before: dict, after: dict) -> set:
    """何かしら変わったシート名の集合（セル・結合・列幅・行高・追加/削除）。"""
    changed = set()
    for name in set(before["sheets"]) | set(after["sheets"]):
        if name not in before["sheets"] or name not in after["sheets"]:
            changed.add(name)
            continue
        if before.get("merges", {}).get(name) != after.get("merges", {}).get(name):
            changed.add(name)
        elif before.get("colw", {}).get(name) != after.get("colw", {}).get(name):
            changed.add(name)
        elif before.get("rowh", {}).get(name) != after.get("rowh", {}).get(name):
            changed.add(name)
    for sheet, _r, _c in _changed_cells(before, after):
        changed.add(sheet)
    return changed


def mention_overlap_advisory(mentions: dict, before: dict, after: dict) -> list:
    """言及があるのに変更範囲と全く重ならない場合だけ警告する（保守的）。
       数字表記の列は 0 起点/1 起点の両解釈を許し、どちらかが触られていれば沈黙する。"""
    if not (mentions["cols"] or mentions.get("digit_cols") or mentions["rows"]
            or mentions["sheets"]):
        return []
    changed = _changed_cells(before, after)
    changed_cols = {c for _, _, c in changed}
    changed_rows = {r for _, r, _ in changed}
    changed_sheets = _changed_sheets(before, after)

    lines = []
    for col in sorted(mentions["cols"]):
        if col not in changed_cols:
            letter = get_column_letter(col)
            lines.append(f"★ 依頼で言及された『列{letter}』は存在しません/変更されていません")
    for n in sorted(mentions.get("digit_cols", set())):
        # 1 起点読み = 列 n / 0 起点読み = 列 n+1 (1 起点換算)。両方外れた時だけ警告し、
        # 警告文は文字に変換せずユーザーの書いた数字のまま返す (推定で上書きしない)
        candidates = {c for c in (n, n + 1) if c >= 1}
        if candidates and not (candidates & changed_cols):
            lines.append(f"★ 依頼で言及された『列{n}』は存在しません/変更されていません")
    for row in sorted(mentions["rows"]):
        if row not in changed_rows:
            lines.append(f"★ 依頼で言及された『行{row}』は存在しません/変更されていません")
    for sheet in sorted(mentions["sheets"]):
        if sheet not in changed_sheets:
            lines.append(f"★ 依頼で言及された『{sheet}』は存在しません/変更されていません")
    return lines


def build_advisories(task: str, before: dict, after: dict) -> list:
    """diff の後に表示する助言行を全部集める。
       ①幽霊データ ②一様埋め ③件数の突き合わせ ④依頼文言との重なり。"""
    lines = []
    for fn in (detect_ghost_data, detect_uniform_fill):
        msg = fn(before, after)
        if msg:
            lines.append(msg)
    recon = count_reconciliation(before, after)
    if recon:
        lines.append(recon)
    mentions = extract_task_mentions(task, before["sheets"])
    lines.extend(mention_overlap_advisory(mentions, before, after))
    return lines


# ---------------------------------------------------------------------------
# M2b: 中間命令言語（DSL）パイプライン
#   ①翻訳(LLM) → ②検証(接地) → ③確認行 → ④決定論 codegen(LLM不使用) → ⑤適用(既存機構)
#   → ⑥op別事後条件(openpyxl で機械検証)
#   ★ 天井は bench/translation_battery.json で実測（凍結合格線: op90%/slot80%/誤断定20%）。
#   CLARIFY・語彙外(FREEFORM)・翻訳失敗は現行の自由生成経路（M2a 助言つき）へ retreat する。
# ---------------------------------------------------------------------------

OP_LABELS = {
    "SORT": "並べ替え", "COMPUTE_COLUMN": "計算列", "LOOKUP_FILL": "転記",
    "AGGREGATE": "集計", "BOLD": "太字", "FILL_COLOR": "背景色",
    "NUMBER_FORMAT": "数値書式", "MERGE": "セル結合", "CHART": "グラフ",
    "CENTER_ALIGN": "中央揃え",
}

# op → 必須 slot 名のタプル。翻訳直後の slot 欠落チェックと確認行の項目順を兼ねる。
OP_SCHEMA = {
    "SORT": ("col", "order"),
    "COMPUTE_COLUMN": ("operands", "operator"),
    "LOOKUP_FILL": ("target_sheet", "target_col", "source_sheet", "key_col"),
    "AGGREGATE": ("group_col", "value_col"),
    "BOLD": ("target",),
    "FILL_COLOR": ("target", "color"),
    "NUMBER_FORMAT": ("col", "style"),
    "MERGE": ("range",),
    "CHART": ("value_col",),
    "CENTER_ALIGN": ("target",),
}

# ★ bench/translation_spike.py（実測 v1）と同じ語彙定義（bench 側は比較用に据え置き、
#   本番プロンプトはここが唯一の元）。
OPS_DOC = """SORT: 並べ替え。args: col(列名), order(asc|desc)
COMPUTE_COLUMN: 既存列同士の計算。args: operands(列名2つ), operator(+,-,*,/), target(省略可・実在する列名。
  依頼が「〜に」のように既存列を名指ししたらその列名を入れる。無指定なら新しい列を作る)
LOOKUP_FILL: 別シートの対応表から値を転記。args: target_sheet, target_col, source_sheet, key_col
AGGREGATE: グループ別に集計表を作る。args: group_col, value_col
BOLD: 太字。args: target("row:行番号" か "col:列名")
FILL_COLOR: 背景色。args: target("row:N"か"col:列名"), color(英語色名)
NUMBER_FORMAT: 数値書式。args: col(列名), style("thousands")
MERGE: セル結合。args: range("A1:C1"形式)
CHART: 棒グラフ。args: value_col(列名)
CENTER_ALIGN: 中央揃え。args: target("all" か "col:列名")"""

# ★ M2c: battery(v1) が実測で取り違えた基本パターンに加え、複合依頼(battery v2)を few-shot で
#   教える。同じ混同/構造を別の言い回しで示す（battery の項目文そのままは使わない＝暗記でなく
#   汎化を確かめる）。
#   ①「引いてくる/転記」は LOOKUP_FILL であって COMPUTE_COLUMN（四則演算）ではない。
#   ②③ 条件付き書式・集計行の追記は語彙に無い操作＝ OUT_OF_VOCAB（曖昧ではないので CLARIFY ではない）。
#   ④ 語彙内(target 付き COMPUTE_COLUMN)＋語彙外の混在。⑤ 語彙内どうしの2連(依存なし)。
TRANSLATION_FEWSHOT = [
    ('対象ブックの構成: {"Sheet": ["商品", "単価"], "商品マスタ": ["商品", "単価"]}\n'
     '依頼: 「商品マスタから商品名を引っ張ってきて明細に入れて」',
     '{"plan": [{"op": "LOOKUP_FILL", "args": {"target_sheet": "Sheet", "target_col": "商品", '
     '"source_sheet": "商品マスタ", "key_col": "商品"}}]}'),
    ('対象ブックの構成: {"Sheet": ["商品", "金額", "在庫"]}\n'
     '依頼: 「金額が1000円未満の行を薄い黄色にして」',
     '{"plan": [{"op": "OUT_OF_VOCAB", "about": "条件付き書式"}]}'),
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「合計を一番下の行に追加して」',
     '{"plan": [{"op": "OUT_OF_VOCAB", "about": "集計行の追記"}]}'),
    ('対象ブックの構成: {"Sheet": ["商品", "数量", "単価", "金額"]}\n'
     '依頼: 「金額に数量×単価を入れて、割引後の金額も出して」',
     '{"plan": ['
     '{"op": "COMPUTE_COLUMN", "args": {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}}, '
     '{"op": "OUT_OF_VOCAB", "about": "割引後の金額"}]}'),
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「金額で昇順に並べ替えて、見出し行を太字にして」',
     '{"plan": ['
     '{"op": "SORT", "args": {"col": "金額", "order": "asc"}}, '
     '{"op": "BOLD", "args": {"target": "row:1"}}]}'),
]

TRANSLATION_SYSTEM = """あなたは表計算操作の翻訳係。日本語の依頼を、下の操作語彙を使った「計画」の JSON に翻訳する。
出力形式は必ず {{"plan": [ {{...}}, {{...}}, ... ]}}。それ以外は書かない。
依頼が複数の操作を含むなら、その全部を計画に順番どおり列挙すること。一部を省略してはいけない
（黙って落とすことを禁止・単一の依頼なら要素数1の計画にする）。

計画の各要素は次のどれか一つ:
- {{"op": "<語彙>", "args": {{...}}}}  操作語彙のどれかに当てはまる場合
- {{"op": "OUT_OF_VOCAB", "about": "<何についての依頼か、短く>"}}  操作語彙のどれにも当てはまらない部分
  （条件付き書式・行/シートの削除やコピー・集計行の追記などは語彙外。必ずこの形で計画に残す）
- {{"op": "CLARIFY", "question": "確認文"}}  依頼が曖昧で必須引数を確定できない場合。推測で断定しない

重要な規則:
- 列は必ず「対象ブックの構成」に実在する列名で指定する（番号ではなく）。ただし直前の段が新規作成する
  列を後続の段が参照する場合は、依頼文の言い方のままでよい（実行時に解決する）
- 依頼が既存の列を名指し（「小計に」等）して値を入れる/書き換える場合、COMPUTE_COLUMN の args に
  target(その実在列名) を入れる
- JSON のみ出力（説明・markdown 柵は禁止）

操作語彙:
{ops}"""

TRANSLATION_USER = "対象ブックの構成: {book}\n依頼: 「{text}」"


def build_translation_messages(task: str, book_meta: dict) -> list:
    """翻訳リクエストの messages（system + few-shot + 実クエリ）を組む。"""
    msgs = [{"role": "system", "content": TRANSLATION_SYSTEM.format(ops=OPS_DOC)}]
    for user_ex, assistant_ex in TRANSLATION_FEWSHOT:
        msgs.append({"role": "user", "content": user_ex})
        msgs.append({"role": "assistant", "content": assistant_ex})
    book_desc = json.dumps(book_meta.get("headers", {}), ensure_ascii=False)
    msgs.append({"role": "user", "content": TRANSLATION_USER.format(book=book_desc, text=task)})
    return msgs


def ollama_generate_json(model: str, messages: list, temperature: float = 0.1,
                          num_predict: int = 300) -> str:
    """M2b 翻訳層専用。/api/chat を format="json" 付きで叩く（bench/translation_spike.py
       で実測済みの型。qwen3 系は think:false を足す）。★ 通常生成の ollama_generate とは
       あえて別関数にする — format="json" を全体に効かせると Basic 生成側が壊れるため。
       戻り値は content 文字列（json.loads は呼び出し側 translate_task が行う）。"""
    body = {
        "model": model, "messages": messages, "stream": False, "format": "json",
        "options": {"temperature": temperature, "num_predict": num_predict, "num_ctx": 8192},
    }
    if "qwen3" in model:
        body["think"] = False
    req = urllib.request.Request(f"{OLLAMA}/api/chat", data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        d = json.load(r)
    return d.get("message", {}).get("content", "")


def _normalize_plan_step(data) -> dict:
    """計画の1要素を正規化する（① 翻訳の per-step 版・translate_task の旧 per-op ロジックを継承）。
       ★ 黙って落とさない: 不明な形・語彙外・必須 slot 欠落は必ず何らかの op を持つ dict として
       残す（呼び出し側がその段を FREEFORM(自由生成)/OUT_OF_VOCAB として扱えるように）。
       - モデルが明示した op="OUT_OF_VOCAB" は about を保って素通し。
       - op が語彙外・必須 slot 欠落・非 dict 要素は op="FREEFORM" にする
        （旧 translate_task の『退避は FREEFORM』という分類をそのまま踏襲。単一依頼のときの
        後方互換＝cmd_run の全面自由生成 retreat が変わらないようにするため）。"""
    if not isinstance(data, dict):
        return {"op": "FREEFORM", "args": {}}
    op = str(data.get("op", "")).upper()
    if op == "CLARIFY":
        question = data.get("question")
        return {"op": "CLARIFY", "question": question or "確認が必要です", "args": {}}
    if op == "OUT_OF_VOCAB":
        about = data.get("about")
        return {"op": "OUT_OF_VOCAB", "about": str(about) if about else "内容不明の依頼", "args": {}}
    if op not in OP_SCHEMA:
        return {"op": "FREEFORM", "args": {}}
    args = data.get("args")
    if not isinstance(args, dict):
        # モデルが args で包まず op と slot をフラットに返した場合の救済（寛容に受ける）。
        args = {k: v for k, v in data.items() if k not in ("op", "about", "question")}
    required = OP_SCHEMA[op]
    if any(k not in args or args[k] in (None, "") for k in required):
        return {"op": "FREEFORM", "args": {}}
    return {"op": op, "args": args}


def translate_task(model: str, task: str, book_meta: dict, temperature: float = 0.1) -> dict:
    """① 翻訳。自然言語タスクを命令言語の「計画」 {"plan": [step, ...]} に翻訳する（M2c）。
       接地: book_meta の実在シート/列名だけを few-shot と一緒に渡す。
       ★ 複合依頼（複数の操作を含む依頼）は plan に全段を列挙する。★ 後方互換: 単一依頼は
       長さ1の計画になる（モデルが "plan" で包まず1操作だけ返した場合もここで長さ1に正規化する）。
       各 step は {"op": <語彙>, "args": {...}} / {"op": "OUT_OF_VOCAB", "about": ...} /
       {"op": "CLARIFY", "question": ...} / {"op": "FREEFORM"}（退避）のいずれか。
       ★ 失敗（API 不通・JSON 不正・空応答）は例外を投げずクラッシュさせない。
       op="FREEFORM" 一段の計画に退避し、呼び出し側が現行の自由生成経路（M2a 助言つき）
       へフォールバックできるようにする。"""
    try:
        messages = build_translation_messages(task, book_meta)
        raw = ollama_generate_json(model, messages, temperature=temperature, num_predict=700)
        data = json.loads(raw)
    except Exception:
        return {"plan": [{"op": "FREEFORM", "args": {}}]}
    steps_raw = None
    if isinstance(data, dict) and isinstance(data.get("plan"), list):
        steps_raw = data["plan"]
    elif isinstance(data, dict):
        steps_raw = [data]   # 後方互換: モデルが plan で包まず単一 op を直接返した
    elif isinstance(data, list):
        steps_raw = data
    if not steps_raw:
        return {"plan": [{"op": "FREEFORM", "args": {}}]}
    return {"plan": [_normalize_plan_step(s) for s in steps_raw]}


# --- ② 検証（接地：実在するシート/列名かを機械照合） -------------------------

COLOR_MAP = {
    "red": "FF0000", "green": "00B050", "blue": "0000FF", "yellow": "FFFF00",
    "orange": "FFA500", "purple": "800080", "pink": "FFC0CB", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "lightyellow": "FFFFE0",
    "lightred": "FFCCCC", "lightgray": "D3D3D3", "lightgrey": "D3D3D3",
}


def _digit_candidates(raw: str, headers: list) -> list:
    """数字表記の列参照を 0 起点/1 起点の両解釈で実在列名に変換した候補（重複除去・順序維持）。
       数字表記でなければ空リスト。"""
    s = str(raw).strip()
    if not re.fullmatch(r"\d+", s):
        return []
    n = int(s)
    cands = []
    for idx in (n, n - 1):        # 0起点読み(n) と 1起点読み(n-1) の両方を試す
        if 0 <= idx < len(headers) and headers[idx] not in cands:
            cands.append(headers[idx])
    return cands


def resolve_col_ref(raw, headers: list) -> tuple:
    """(実在列名 or None, 推定だったか, エラー文 or None)。
       ★ 列名を正とする。実在すればそのまま。数字表記なら 0/1 起点の両候補を試し、
       一意に決まればそれを『推定』として解決、決まらなければ CLARIFY 相当のエラーを返す。"""
    s = str(raw)
    if s in headers:
        return s, False, None
    cands = _digit_candidates(s, headers)
    if len(cands) == 1:
        return cands[0], True, None
    if len(cands) > 1:
        return None, False, f"列『{s}』は複数の解釈が可能で一意に決まりません: {cands}"
    known = ", ".join(headers) if headers else "(無し)"
    return None, False, f"列『{s}』がありません。ある列: {known}"


def verify_dsl_args(op: str, args: dict, book_meta: dict) -> tuple:
    """② 検証。(ok, resolved_args, inferred_keys, error_message)。
       args のシート/列名が実在するかを機械照合し、実在名に解決する。実在しなければ
       CLARIFY 相当のエラーメッセージを返す（呼び出し側が確認質問として表示する）。"""
    sheets = book_meta["sheets"]
    headers = book_meta["headers"]
    if not sheets:
        return False, dict(args), set(), "ブックにシートが無い"
    first_sheet = sheets[0]
    resolved = dict(args)
    inferred: set = set()

    def resolve_in(key: str, sheet_name: str):
        val, was_inferred, err = resolve_col_ref(resolved.get(key), headers.get(sheet_name, []))
        if err:
            return err
        resolved[key] = val
        if was_inferred:
            inferred.add(key)
        return None

    def check_sheet(key: str):
        name = resolved.get(key)
        if name not in sheets:
            return f"シート『{name}』がありません。あるシート: {', '.join(sheets)}"
        return None

    if op == "SORT":
        if (err := resolve_in("col", first_sheet)):
            return False, resolved, inferred, err
        if resolved.get("order") not in ("asc", "desc"):
            return False, resolved, inferred, f"順序『{resolved.get('order')}』は asc/desc のどちらでもありません"

    elif op == "COMPUTE_COLUMN":
        operands = resolved.get("operands")
        if not (isinstance(operands, list) and len(operands) == 2):
            return False, resolved, inferred, "演算対象が2つの列名になっていません"
        new_operands = []
        for o in operands:
            v, was_inferred, err = resolve_col_ref(o, headers.get(first_sheet, []))
            if err:
                return False, resolved, inferred, err
            new_operands.append(v)
            if was_inferred:
                inferred.add("operands")
        resolved["operands"] = new_operands
        if resolved.get("operator") not in ("+", "-", "*", "/"):
            return False, resolved, inferred, f"演算子『{resolved.get('operator')}』が不明です"
        # ★ M2c: target(任意) — 依頼が既存列を名指し（「小計に」等）した場合はその列に書く。
        #   無指定なら従来どおり新規列（codegen_dsl 側で分岐）。
        if resolved.get("target"):
            v, was_inferred, err = resolve_col_ref(resolved["target"], headers.get(first_sheet, []))
            if err:
                return False, resolved, inferred, err
            resolved["target"] = v
            if was_inferred:
                inferred.add("target")

    elif op == "LOOKUP_FILL":
        if (err := check_sheet("target_sheet")):
            return False, resolved, inferred, err
        if (err := check_sheet("source_sheet")):
            return False, resolved, inferred, err
        if resolved["target_sheet"] != first_sheet:
            return False, resolved, inferred, f"対象シートは1枚目（{first_sheet}）のみ対応しています"
        if (err := resolve_in("target_col", resolved["target_sheet"])):
            return False, resolved, inferred, err
        if (err := resolve_in("key_col", resolved["target_sheet"])):
            return False, resolved, inferred, err

    elif op == "AGGREGATE":
        if (err := resolve_in("group_col", first_sheet)):
            return False, resolved, inferred, err
        if (err := resolve_in("value_col", first_sheet)):
            return False, resolved, inferred, err

    elif op in ("BOLD", "FILL_COLOR", "CENTER_ALIGN"):
        target = str(resolved.get("target", ""))
        if target == "all":
            if op != "CENTER_ALIGN":
                return False, resolved, inferred, f"対象『all』は {OP_LABELS[op]} では未対応です"
        elif target.startswith("row:"):
            n = target[4:]
            if not (n.isdigit() and int(n) >= 1):
                return False, resolved, inferred, f"行番号『{n}』が不正です"
        elif target.startswith("col:"):
            colname = target[4:]
            v, was_inferred, err = resolve_col_ref(colname, headers.get(first_sheet, []))
            if err:
                return False, resolved, inferred, err
            resolved["target"] = f"col:{v}"
            if was_inferred:
                inferred.add("target")
        else:
            return False, resolved, inferred, f"対象『{target}』の形式が不明です（row:N / col:列名 / all）"
        if op == "FILL_COLOR":
            color = str(resolved.get("color", "")).lower()
            if color not in COLOR_MAP:
                return False, resolved, inferred, f"色『{color}』は未対応です。使える色: {', '.join(sorted(COLOR_MAP))}"
            resolved["color"] = color

    elif op == "NUMBER_FORMAT":
        if (err := resolve_in("col", first_sheet)):
            return False, resolved, inferred, err
        if resolved.get("style") != "thousands":
            return False, resolved, inferred, f"書式『{resolved.get('style')}』は未対応です（対応: thousands）"

    elif op == "MERGE":
        if not re.fullmatch(r"[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+", str(resolved.get("range", ""))):
            return False, resolved, inferred, f"範囲『{resolved.get('range')}』の形式が不正です（例: A1:C1）"

    elif op == "CHART":
        if (err := resolve_in("value_col", first_sheet)):
            return False, resolved, inferred, err

    else:
        return False, resolved, inferred, f"未対応の操作: {op}"

    return True, resolved, inferred, None


# --- ③ 確認行（命令言語形式） -------------------------------------------------

_CONFIRM_FIELDS = {
    "SORT": (("対象", "col", None), ("順", "order", lambda v: "降順" if v == "desc" else "昇順")),
    "COMPUTE_COLUMN": (("演算対象", "operands", lambda v: " と ".join(v)), ("演算子", "operator", None),
                        ("対象列", "target", None)),
    "LOOKUP_FILL": (("対象シート", "target_sheet", None), ("対象列", "target_col", None),
                     ("参照シート", "source_sheet", None), ("キー列", "key_col", None)),
    "AGGREGATE": (("分類列", "group_col", None), ("集計列", "value_col", None)),
    "BOLD": (("対象", "target", None),),
    "FILL_COLOR": (("対象", "target", None), ("色", "color", None)),
    "NUMBER_FORMAT": (("対象列", "col", None), ("書式", "style", None)),
    "MERGE": (("範囲", "range", None),),
    "CHART": (("値列", "value_col", None),),
    "CENTER_ALIGN": (("対象", "target", None),),
}


def format_confirmation_line(op: str, resolved_args: dict, inferred: set) -> str:
    """命令言語形式の確認行を1行で組む（例: 解釈: 操作:並べ替え 対象:金額 順:降順）。
       推定で埋めた（数字表記から解決した等）引数には (推定) を付ける。
       ★ M2c: キー自体が resolved_args に無い任意項目（COMPUTE_COLUMN の target 等）は
       そのフィールドを丸ごと省略する（必須項目は常に存在するので既存の表示は変わらない）。"""
    parts = [f"操作:{OP_LABELS.get(op, op)}"]
    for label, key, transform in _CONFIRM_FIELDS.get(op, ()):
        if key not in resolved_args:
            continue
        val = resolved_args.get(key)
        shown = transform(val) if transform else val
        tag = "(推定)" if key in inferred else ""
        parts.append(f"{label}:{shown}{tag}")
    return "解釈: " + " ".join(parts)


# --- ④ 決定論 codegen（op → Basic テンプレ。LLM は一切使わない） --------------

def _wrap_basic(body: str) -> str:
    """CONTRACT と同じ骨格（Option 2行 + Sub Run(oDoc As Object) 1つ）で包む。"""
    return "Option VBASupport 1\nOption Explicit\n\nSub Run(oDoc As Object)\n" + body + "End Sub\n"


def _scan_last_row_basic(var: str = "oSheet", key_col: str = "0") -> str:
    """走査ループの定型（refs の作法どおり：A列を上から走査して最終データ行を探す）。"""
    return (f"    lastRow = 1\n"
            f"    Do While {var}.getCellByPosition({key_col}, lastRow).getString() <> \"\"\n"
            f"        lastRow = lastRow + 1\n"
            f"    Loop\n"
            f"    lastRow = lastRow - 1\n"
            f"    If lastRow < 1 Then Exit Sub\n")


def codegen_dsl(op: str, resolved_args: dict, book_meta: dict) -> str:
    """④ 決定論 codegen。既存ヘルパへの Call を最優先し、無い操作だけテンプレ Basic を書く。"""
    headers = book_meta["headers"]
    first_sheet = book_meta["sheets"][0]

    if op == "SORT":
        col_idx = headers[first_sheet].index(resolved_args["col"])
        asc = "True" if resolved_args["order"] == "asc" else "False"
        return _wrap_basic(f"    Call SortByColumn(oDoc, {col_idx}, {asc})\n")

    if op == "LOOKUP_FILL":
        theaders = headers[resolved_args["target_sheet"]]
        key_idx = theaders.index(resolved_args["key_col"])
        tgt_idx = theaders.index(resolved_args["target_col"])
        src = resolved_args["source_sheet"].replace('"', '""')
        return _wrap_basic(f'    Call VLookupFromTable(oDoc, {key_idx}, {tgt_idx}, "{src}")\n')

    if op == "AGGREGATE":
        g_idx = headers[first_sheet].index(resolved_args["group_col"])
        v_idx = headers[first_sheet].index(resolved_args["value_col"])
        return _wrap_basic(f"    Call SummaryTable(oDoc, {g_idx}, {v_idx})\n")

    if op == "NUMBER_FORMAT":
        col_idx = headers[first_sheet].index(resolved_args["col"])
        return _wrap_basic(f"    Call FormatThousands(oDoc, {col_idx})\n")

    if op == "MERGE":
        c1s, r1s, c2s, r2s = re.match(
            r"([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)", resolved_args["range"]).groups()
        col1 = column_index_from_string(c1s.upper()) - 1
        col2 = column_index_from_string(c2s.upper()) - 1
        row1, row2 = int(r1s) - 1, int(r2s) - 1
        return _wrap_basic(f"    Call MergeCells(oDoc, {col1}, {row1}, {col2}, {row2})\n")

    if op == "CHART":
        v_idx = headers[first_sheet].index(resolved_args["value_col"])
        return _wrap_basic(f"    Call InsertBarChart(oDoc, {v_idx})\n")

    if op == "CENTER_ALIGN":
        if resolved_args["target"] == "all":
            return _wrap_basic("    Call AlignCenter(oDoc)\n")
        # col:NAME はヘルパ無し → refs の作法（走査して範囲を求め HoriJustify）でテンプレを書く。
        col_idx = headers[first_sheet].index(resolved_args["target"][4:])
        body = ("    Dim oSheet As Object, oRange As Object, lastRow As Long\n"
                "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                + _scan_last_row_basic().replace("lastRow < 1", "lastRow < 0")
                + f"    oRange = oSheet.getCellRangeByPosition({col_idx}, 0, {col_idx}, lastRow)\n"
                "    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER\n")
        return _wrap_basic(body)

    if op == "BOLD":
        target = resolved_args["target"]
        if target.startswith("row:"):
            row_idx = int(target[4:]) - 1
            body = ("    Dim oSheet As Object, lastCol As Integer\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    "    lastCol = 0\n"
                    "    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> \"\"\n"
                    "        lastCol = lastCol + 1\n"
                    "    Loop\n"
                    "    lastCol = lastCol - 1\n"
                    "    If lastCol < 0 Then Exit Sub\n"
                    f"    Call StyleBold(oDoc, 0, {row_idx}, lastCol, {row_idx})\n")
        else:
            col_idx = headers[first_sheet].index(target[4:])
            body = ("    Dim oSheet As Object, lastRow As Long\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    + _scan_last_row_basic().replace("lastRow < 1", "lastRow < 0")
                    + f"    Call StyleBold(oDoc, {col_idx}, 0, {col_idx}, lastRow)\n")
        return _wrap_basic(body)

    if op == "FILL_COLOR":
        target = resolved_args["target"]
        hexcolor = COLOR_MAP[resolved_args["color"]]
        if target.startswith("row:"):
            row_idx = int(target[4:]) - 1
            body = ("    Dim oSheet As Object, lastCol As Integer, c As Integer\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    "    lastCol = 0\n"
                    "    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> \"\"\n"
                    "        lastCol = lastCol + 1\n"
                    "    Loop\n"
                    "    lastCol = lastCol - 1\n"
                    "    If lastCol < 0 Then Exit Sub\n"
                    "    For c = 0 To lastCol\n"
                    f"        oSheet.getCellByPosition(c, {row_idx}).CellBackColor = &H{hexcolor}&\n"
                    "    Next c\n")
        else:
            col_idx = headers[first_sheet].index(target[4:])
            body = ("    Dim oSheet As Object, lastRow As Long, r As Long\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    + _scan_last_row_basic().replace("lastRow < 1", "lastRow < 0")
                    + "    For r = 0 To lastRow\n"
                    f"        oSheet.getCellByPosition({col_idx}, r).CellBackColor = &H{hexcolor}&\n"
                    "    Next r\n")
        return _wrap_basic(body)

    if op == "COMPUTE_COLUMN":
        op1, op2 = resolved_args["operands"]
        i1 = headers[first_sheet].index(op1)
        i2 = headers[first_sheet].index(op2)
        operator = resolved_args["operator"]
        target = resolved_args.get("target")
        # ★ M2c: target(実在列名) 指定時はその列に書く（新規列を作らない）。
        #   無指定なら従来どおり次の空き列に新規列を作る。
        if target:
            new_col = headers[first_sheet].index(target)
            header_write = ""   # 既存の見出しはそのまま（上書きしない）
        else:
            new_col = len(headers[first_sheet])   # 0起点で次の空き列
            header_name = f"{op1}{operator}{op2}".replace('"', '""')
            header_write = f"    oSheet.getCellByPosition({new_col}, 0).setString(\"{header_name}\")\n"
        body = ("    Dim oSheet As Object, lastRow As Long, i As Long\n"
                "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                + _scan_last_row_basic()
                + header_write
                + "    For i = 1 To lastRow\n"
                f"        oSheet.getCellByPosition({new_col}, i).setValue("
                f"oSheet.getCellByPosition({i1}, i).getValue() {operator} "
                f"oSheet.getCellByPosition({i2}, i).getValue())\n"
                "    Next i\n")
        return _wrap_basic(body)

    raise ValueError(f"未対応の op: {op}")


# --- ⑥ op 別事後条件（達成の機械検証。openpyxl で out ファイルを読むだけ・LO 不要） ----

def _col_index_by_header(ws, name: str):
    """見出し行(1行目)を左から走査して name に一致する列の1起点インデックスを返す。無ければ None。"""
    c = 1
    while True:
        v = ws.cell(row=1, column=c).value
        if v in (None, ""):
            return None
        if str(v) == name:
            return c
        c += 1


def _scan_last_row(ws, key_col: int = 1) -> int:
    """key_col(1起点)を上から走査した最終データ行（見出し行0を除く）。データが無ければ1。"""
    r = 2
    while ws.cell(row=r, column=key_col).value not in (None, ""):
        r += 1
    return r - 1


def _scan_last_col(ws) -> int:
    """見出し行(1行目)を左から走査した最終列（1起点）。"""
    c = 1
    while ws.cell(row=1, column=c).value not in (None, ""):
        c += 1
    return c - 1


def _apply_operator(a, b, operator: str):
    a = a or 0
    b = b or 0
    if operator == "+":
        return a + b
    if operator == "-":
        return a - b
    if operator == "*":
        return a * b
    if operator == "/":
        return a / b if b else 0
    raise ValueError(operator)


def check_sort(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    idx = _col_index_by_header(ws, args["col"])
    if idx is None:
        wb.close()
        return False, f"列『{args['col']}』が見つからない"
    last = _scan_last_row(ws)
    vals = [ws.cell(row=r, column=idx).value for r in range(2, last + 1)]
    wb.close()
    if len(vals) < 2:
        return True, "行数が少なく比較不要"
    asc = args["order"] == "asc"
    ok = (all(vals[i] <= vals[i + 1] for i in range(len(vals) - 1)) if asc
          else all(vals[i] >= vals[i + 1] for i in range(len(vals) - 1)))
    if not ok:
        return False, f"列『{args['col']}』が指定順（{args['order']}）に並んでいない"
    return True, f"{len(vals)} 行が{'昇順' if asc else '降順'}"


def check_compute_column(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    op1, op2 = args["operands"]
    i1 = _col_index_by_header(ws, op1)
    i2 = _col_index_by_header(ws, op2)
    # ★ M2c: target(実在列名) 指定時はその列を検証する。無指定なら従来どおり自動命名の新列。
    target = args.get("target")
    newname = target or f"{op1}{args['operator']}{op2}"
    inew = _col_index_by_header(ws, newname)
    if i1 is None or i2 is None or inew is None:
        wb.close()
        return False, f"演算対象または対象列『{newname}』が見つからない"
    last = _scan_last_row(ws)
    for r in range(2, last + 1):
        a = ws.cell(row=r, column=i1).value
        b = ws.cell(row=r, column=i2).value
        got = ws.cell(row=r, column=inew).value
        want = _apply_operator(a, b, args["operator"])
        if got is None or abs(got - want) > 1e-6:
            wb.close()
            return False, f"{r}行目: 期待 {want} 実際 {got}"
    wb.close()
    return True, f"{last - 1} 行を検証"


def check_lookup_fill(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    if args["target_sheet"] not in wb.sheetnames or args["source_sheet"] not in wb.sheetnames:
        wb.close()
        return False, "対象/参照シートが無い"
    tws = wb[args["target_sheet"]]
    sws = wb[args["source_sheet"]]
    key_idx = _col_index_by_header(tws, args["key_col"])
    tgt_idx = _col_index_by_header(tws, args["target_col"])
    if key_idx is None or tgt_idx is None:
        wb.close()
        return False, "対象シートにキー列/対象列が無い"
    lookup = {}
    r = 2
    while sws.cell(row=r, column=1).value not in (None, ""):
        lookup[sws.cell(row=r, column=1).value] = sws.cell(row=r, column=2).value
        r += 1
    checked = 0
    r = 2
    while tws.cell(row=r, column=key_idx).value not in (None, ""):
        key = tws.cell(row=r, column=key_idx).value
        if key in lookup:
            got = tws.cell(row=r, column=tgt_idx).value
            want = lookup[key]
            if got != want:
                wb.close()
                return False, f"{r}行目: キー『{key}』の転記値が不一致 (期待 {want!r} 実際 {got!r})"
            checked += 1
        r += 1
    wb.close()
    if checked == 0:
        return False, "対応表に載っているキーが1件も転記されていない"
    return True, f"{checked} 行を検証"


def check_aggregate(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    if "集計" not in wb.sheetnames:
        wb.close()
        return False, "『集計』シートが無い"
    src = wb[wb.sheetnames[0]]
    gi = _col_index_by_header(src, args["group_col"])
    vi = _col_index_by_header(src, args["value_col"])
    if gi is None or vi is None:
        wb.close()
        return False, "分類列/集計列が見つからない"
    expect: dict = {}
    r = 2
    while src.cell(row=r, column=1).value not in (None, ""):
        k = src.cell(row=r, column=gi).value
        v = src.cell(row=r, column=vi).value or 0
        expect[k] = expect.get(k, 0) + v
        r += 1
    out = wb["集計"]
    seen = set()
    r = 2
    while True:
        k = out.cell(row=r, column=1).value
        if k in (None, "") or k == "合計":
            break
        v = out.cell(row=r, column=2).value or 0
        if k not in expect or abs(v - expect[k]) > 1e-6:
            wb.close()
            return False, f"グループ『{k}』の合計が不一致 (期待 {expect.get(k)} 実際 {v})"
        seen.add(k)
        r += 1
    wb.close()
    if seen != set(expect.keys()):
        return False, "集計に含まれないグループがある"
    return True, f"{len(expect)} グループを検証"


def check_bold(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    kind, val = args["target"].split(":", 1)
    if kind == "row":
        last_col = _scan_last_col(ws)
        row = int(val)
        cells = [ws.cell(row=row, column=c) for c in range(1, last_col + 1)]
        label = f"{row}行目"
    else:
        idx = _col_index_by_header(ws, val)
        if idx is None:
            wb.close()
            return False, f"列『{val}』が見つからない"
        last_row = _scan_last_row(ws)
        cells = [ws.cell(row=r, column=idx) for r in range(1, last_row + 1)]
        label = f"列『{val}』"
    ok = all(c.font and c.font.bold for c in cells)
    wb.close()
    if not ok:
        return False, f"{label} に太字でないセルがある"
    return True, f"{len(cells)} セルが太字"


def check_fill_color(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    want_hex = COLOR_MAP[args["color"]].upper()
    kind, val = args["target"].split(":", 1)
    if kind == "row":
        last_col = _scan_last_col(ws)
        row = int(val)
        cells = [ws.cell(row=row, column=c) for c in range(1, last_col + 1)]
        label = f"{row}行目"
    else:
        idx = _col_index_by_header(ws, val)
        if idx is None:
            wb.close()
            return False, f"列『{val}』が見つからない"
        last_row = _scan_last_row(ws)
        cells = [ws.cell(row=r, column=idx) for r in range(1, last_row + 1)]
        label = f"列『{val}』"

    def _matches(cell) -> bool:
        if cell.fill is None or not cell.fill.patternType:
            return False
        return str(cell.fill.start_color.rgb).upper().endswith(want_hex)

    ok = all(_matches(c) for c in cells)
    wb.close()
    if not ok:
        return False, f"{label} に色『{args['color']}』が付いていないセルがある"
    return True, f"{len(cells)} セルの背景色を確認"


def check_number_format(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    idx = _col_index_by_header(ws, args["col"])
    if idx is None:
        wb.close()
        return False, f"列『{args['col']}』が見つからない"
    last = _scan_last_row(ws)
    if last < 2:
        wb.close()
        return False, "データ行が無い"
    ok = all("#,##0" in (ws.cell(row=r, column=idx).number_format or "") for r in range(2, last + 1))
    wb.close()
    if not ok:
        return False, f"列『{args['col']}』に桁区切り書式が付いていないセルがある"
    return True, f"{last - 1} 行に桁区切り書式を確認"


def check_merge(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    ranges = {str(r) for r in ws.merged_cells.ranges}
    wb.close()
    if args["range"] not in ranges:
        return False, f"範囲『{args['range']}』が結合されていない"
    return True, f"{args['range']} の結合を確認"


def check_chart(path: Path, before_charts: int) -> tuple:
    after = _charts_count(path)
    if after != before_charts + 1:
        return False, f"グラフ数が +1 でない（{before_charts} → {after}）"
    return True, f"グラフ数 {before_charts} → {after}"


def check_center_align(path: Path, args: dict) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    target = args["target"]
    if target == "all":
        last_row = _scan_last_row(ws)
        last_col = _scan_last_col(ws)
        cells = [ws.cell(row=r, column=c) for r in range(1, last_row + 1) for c in range(1, last_col + 1)]
        label = "表全体"
    else:
        colname = target[4:]
        idx = _col_index_by_header(ws, colname)
        if idx is None:
            wb.close()
            return False, f"列『{colname}』が見つからない"
        last_row = _scan_last_row(ws)
        cells = [ws.cell(row=r, column=idx) for r in range(1, last_row + 1)]
        label = f"列『{colname}』"
    ok = all(c.alignment and c.alignment.horizontal == "center" for c in cells)
    wb.close()
    if not ok:
        return False, f"{label} に中央揃えでないセルがある"
    return True, f"{len(cells)} セルの中央揃えを確認"


POSTCONDITIONS = {
    "SORT": check_sort, "COMPUTE_COLUMN": check_compute_column,
    "LOOKUP_FILL": check_lookup_fill, "AGGREGATE": check_aggregate,
    "BOLD": check_bold, "FILL_COLOR": check_fill_color,
    "NUMBER_FORMAT": check_number_format, "MERGE": check_merge,
    "CENTER_ALIGN": check_center_align,
}


def run_postcondition(op: str, out_book: Path, resolved_args: dict, before_charts: int = 0) -> tuple:
    """⑥ op 別事後条件。(ok, reason)。CHART だけ before_charts と比較する専用の形。"""
    if op == "CHART":
        return check_chart(out_book, before_charts)
    fn = POSTCONDITIONS.get(op)
    if fn is None:
        return False, f"未対応の op: {op}"
    return fn(out_book, resolved_args)


# ---------------------------------------------------------------------------
# basrun 経由の適用
# ---------------------------------------------------------------------------

def _timeout_error_message(timeout: float) -> str:
    return f"実行時エラー: マクロが {timeout:.0f} 秒で終了しない（無限ループの可能性）"


def short_error_summary(err: str) -> str:
    """★ M2a: obasync 由来の生 Python トレースバックを端末にそのまま出さないための
       整形。最終行（例外名+メッセージ）だけを取り出す。全文は履歴 jsonl 側に残す
       （修復ループへは従来どおり err の全文を渡す＝モデルへの情報は減らさない）。"""
    lines = [ln for ln in (err or "").splitlines() if ln.strip()]
    return lines[-1].strip() if lines else "(詳細不明)"


def _kill_process_tree(pid: int) -> None:
    """PID 指定でプロセスツリーを kill する。
       ★ taskkill /IM（名前一括）は使わない — 無関係な他プロセスも巻き込む
       （2026-08 の Senior MCP 事故の教訓）。必ず特定した PID だけを狙う。"""
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                        capture_output=True)
    else:
        import signal
        try:
            os.killpg(os.getpgid(pid), signal.SIGKILL)
        except Exception:
            pass


def basrun_apply(book: Path, code: str, workdir: Path, helper_files=(),
                  timeout: float | None = DEFAULT_APPLY_TIMEOUT) -> tuple:
    """生成コードを book に適用する。(ok, error_or_None, raw_output)。
       helper_files があれば同じ src に置く＝同じライブラリに同期され、Gen から呼べる。

    ★ M1: timeout（既定 180 秒、None/0 で無効）を basrun 側の `apply --timeout` に
       転送する。basrun は自分の内部タイムアウトで、ハングした接続先の LibreOffice
       だけを stop_office() で終了させる（taskkill 一括はしない、既存機構）。
       ここではさらに外側の安全網として、basrun 自身が固まった場合に備えて
       ゆとり(+30秒)を持たせた outer timeout を持ち、それも超えたら basrun.py の
       プロセスを PID 指定で kill する。どちらの経路で落ちても、修復ループには
       同じ「実行時エラー: マクロが N 秒で終了しない」で渡す（正直な分類）。"""
    src = workdir / "src"
    if src.exists():
        shutil.rmtree(src)
    src.mkdir(parents=True)
    for hf in helper_files:
        shutil.copy2(hf, src / hf.name)
    (src / "Gen.bas").write_text(code, encoding="utf-8")

    cmd = [sys.executable, str(basrun_path()), "apply", str(book), str(src), "AiLine", "Gen.Run"]
    if timeout:
        cmd += ["--timeout", str(timeout)]
    outer_timeout = (timeout + 30) if timeout else None

    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace")  # ★ cp932 事故を避ける
    try:
        out, err_out = proc.communicate(timeout=outer_timeout)
    except subprocess.TimeoutExpired:
        # ★ basrun 自身の内部タイムアウト＋stop_office() が働かなかった場合の安全網。
        _kill_process_tree(proc.pid)
        try:
            out, err_out = proc.communicate(timeout=5)
        except Exception:
            out, err_out = "", ""
        raw = (out or "") + "\n" + (err_out or "")
        return False, _timeout_error_message(timeout), raw

    raw = (out or "") + "\n" + (err_out or "")
    if proc.returncode != 0:
        # basrun 自身の内部タイムアウトで落ちた場合も、同じ「実行時エラー」分類に正規化する。
        if timeout and "秒応答しなかった" in raw:
            return False, _timeout_error_message(timeout), raw
        return False, raw.strip()[-800:], raw
    return True, None, raw


# ★ 正規化パス専用の空マクロ。何もしない（Call すら書かない）が、basrun_apply の
#   `doc.store()` を一度通すことで LibreOffice 側の初回保存の実体化を先に済ませる。
NOOP_MACRO = "Option VBASupport 1\nOption Explicit\n\nSub Run(oDoc As Object)\nEnd Sub\n"


def _stop_office() -> None:
    """basrun.py stop を呼んで LibreOffice を落とす（M2c: normalize_book の自動リトライ用）。
       ★ ここでの失敗は無視する（次の apply がどのみち再起動を試みる。taskkill 一括はしない
       既存機構の外側なので、ここでは basrun 自身の stop に委譲するだけ）。"""
    try:
        subprocess.run([sys.executable, str(basrun_path()), "stop"],
                        capture_output=True, encoding="utf-8", errors="replace")
    except Exception:
        pass


def normalize_book(book: Path, workdir: Path,
                    timeout: float | None = DEFAULT_APPLY_TIMEOUT) -> Path:
    """コピーを LibreOffice で一度（空マクロで）開いて保存する ＝ P0 の正規化パス。

    LibreOffice は openpyxl 製（＝ LO で保存されたことがない）ブックを初回保存する際、
    行高（時に列幅）を実体化する。before スナップショットをこの実体化の**前**に取ると、
    その副作用が「マクロが変化させた」と誤検出される（no-op ガードの偽陽性・製品の心臓）。
    先にこの正規化を一度済ませておけば、以降の before/after 比較はマクロの実際の効果
    だけを見る。コストは LO 往復 1 回（数秒）— 正しさ優先で受け入れる。
    参考: ailine-ts の tests/e/_harness.ts normalizeThroughLibreOffice が同じ手当てを
    テスト側で先に実装していた（挙動の参考。製品経路に入れるのはこちらが初）。

    ★ M2c: 監査2回で2回再現した既知の摩擦（RuntimeException: Could not create system
    bitmap! 等、LibreOffice 側の一時的な描画/接続不調）への低リスク対処。1回失敗しても
    即座に落とさず、stop（LibreOffice を落とす）→ 再起動を挟んで1回だけ自動リトライする。
    それでも失敗したら現行どおりのエラーで落ちる（無限リトライはしない）。"""
    normalized = workdir / ("normalized" + book.suffix)
    shutil.copy2(book, normalized)
    ok, err, _ = basrun_apply(normalized, NOOP_MACRO, workdir, timeout=timeout)
    if not ok:
        _stop_office()
        shutil.copy2(book, normalized)   # 中途半端な保存状態を残さず作り直す
        ok, err, _ = basrun_apply(normalized, NOOP_MACRO, workdir, timeout=timeout)
    if not ok:
        sys.exit(f"正規化パスに失敗した（LibreOffice で開けなかった）: {err}")
    return normalized


def success_message(result: dict) -> str | None:
    """★ の注意書きは『変化を検出して適用が成功した』ときだけ出す。
       失敗(exit 1)や --dry（何も適用していない）で出すのは不誠実（P1）。"""
    if result.get("ok") and not result.get("dry"):
        return "★ 変化は検出したが『正しいか』は上の差分を見て判断してください（no-op ガードは正しさを保証しない）。"
    return None


# ---------------------------------------------------------------------------
# --inplace バックアップ + restore（M2a）
#
# ★ --inplace は原本を上書きするので、上書き前に必ず ~/.ailine/backups/ へコピーする。
#   バックアップに失敗したら --inplace 自体を中止する（安全側。原本は無変更のまま）。
#   restore は復元前の現状も退避してから上書きする＝復元自体も取り消せる。
# ---------------------------------------------------------------------------

_BACKUP_TS_RE = re.compile(r"^\d{8}T\d{6}Z$")


def _utc_ts() -> str:
    """ファイル名に使える UTC タイムスタンプ（例: 20260814T120000Z）。"""
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def backup_path_for(book: Path, ts: str | None = None) -> Path:
    ts = ts or _utc_ts()
    return BACKUP_DIR / f"{book.stem}.{ts}{book.suffix}"


def prune_backups(book: Path, keep: int = DEFAULT_KEEP_BACKUPS) -> list:
    """★ M2c: book の世代のうち keep 件を超える古いもの（list_backups は新しい順）を削除する。
       戻り値は削除したパスのリスト。keep < 0 は「無制限（削除しない）」扱い。"""
    if keep < 0:
        return []
    backups = list_backups(book)
    stale = backups[keep:]
    deleted = []
    for p in stale:
        try:
            p.unlink()
            deleted.append(p)
        except OSError:
            pass
    return deleted


def make_backup(book: Path, keep: int = DEFAULT_KEEP_BACKUPS) -> Path:
    """book のバックアップを ~/.ailine/backups/ に作る。戻り値はバックアップ先。
       ★ 失敗したら例外を投げる（呼び出し側が --inplace 中止の判断に使う）。
       ★ M2c: 新しいバックアップを作った後、keep 世代を超えた古いものを剪定する
       （既定 DEFAULT_KEEP_BACKUPS=10。無制限にすると個人開発機のディスクを静かに食う）。"""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    dst = backup_path_for(book)
    shutil.copy2(book, dst)
    prune_backups(book, keep=keep)
    return dst


def _parse_backup_name(name: str, stem: str, suffix: str) -> str | None:
    """バックアップのファイル名が `<stem>.<ts><suffix>` の形かを見て、ts を返す
       （形が違えば None）。"""
    prefix = stem + "."
    if not (name.startswith(prefix) and name.endswith(suffix)):
        return None
    ts = name[len(prefix):len(name) - len(suffix)]
    return ts if _BACKUP_TS_RE.match(ts) else None


def list_backups(book: Path) -> list:
    """book に対応するバックアップを新しい順(タイムスタンプ降順)で返す。"""
    if not BACKUP_DIR.is_dir():
        return []
    stem, suffix = book.stem, book.suffix
    found = []
    for p in BACKUP_DIR.iterdir():
        ts = _parse_backup_name(p.name, stem, suffix)
        if ts is not None:
            found.append((ts, p))
    found.sort(key=lambda pair: pair[0], reverse=True)   # ts は辞書順=時刻順
    return [p for _ts, p in found]


def restore_backup(book: Path) -> Path:
    """book を最新バックアップから復元する。★ 復元前の現状も退避してからコピーする
       （復元自体も取り消せるように）。戻り値は使ったバックアップの Path。
       バックアップが1つも無ければ例外を投げる。"""
    backups = list_backups(book)
    if not backups:
        raise FileNotFoundError(f"{book.name} のバックアップが無い")
    latest = backups[0]
    if book.exists():
        make_backup(book)   # 復元前の現状も退避＝restore 自体も可逆にする
    shutil.copy2(latest, book)
    return latest


def cmd_restore(a: argparse.Namespace) -> int:
    book = Path(a.book).resolve()
    if a.list:
        backups = list_backups(book)
        if not backups:
            print(f"{book.name} のバックアップは無い")
            return 0
        print(f"{book.name} のバックアップ（{len(backups)} 世代・新しい順）:")
        for p in backups:
            print(p.name)
        return 0
    try:
        used = restore_backup(book)
    except FileNotFoundError as e:
        print(f"× {e}")
        return 1
    print(f"✓ {book.name} を {used.name} から復元した")
    return 0


# ---------------------------------------------------------------------------
# doctor（M1: セットアップ診断）
# ---------------------------------------------------------------------------

def _load_module_from_path(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _check_python_version() -> tuple:
    ok = sys.version_info >= (3, 10)
    detail = "" if ok else f"現在 {sys.version.split()[0]}。3.10 以上へ更新して"
    return ok, detail


def _check_openpyxl() -> tuple:
    try:
        import openpyxl as _op  # noqa: F401 — 到達確認のみ
        return True, ""
    except ImportError:
        return False, "pip install openpyxl"


def _check_ollama_reachable(timeout: float = 3.0) -> tuple:
    try:
        with urllib.request.urlopen(f"{OLLAMA}/api/tags", timeout=timeout) as r:
            json.load(r)
        return True, ""
    except Exception as e:
        return False, f"`ollama serve` を起動して（{e}）"


def _check_model_available(model: str, timeout: float = 3.0) -> tuple:
    try:
        with urllib.request.urlopen(f"{OLLAMA}/api/tags", timeout=timeout) as r:
            d = json.load(r)
    except Exception:
        return False, "ollama に繋がらないため確認できない（先に ollama 到達を直して）"
    names = {m.get("name") for m in d.get("models", [])}
    if model in names:
        return True, ""
    return False, f"`ollama pull {model}` で取得して"


def _check_libreoffice() -> tuple:
    p = _find_basrun_path()
    if p is None:
        return False, "basrun.py が無いため確認できない"
    try:
        mod = _load_module_from_path(p, "_ailine_basrun_probe")
        d = mod.office_dir()
        return True, str(d)
    except SystemExit as e:
        return False, str(e)
    except Exception as e:
        return False, f"検出に失敗: {e}"


def _check_basrun() -> tuple:
    p = _find_basrun_path()
    if p is None:
        return False, ("ailine と並びに https://github.com/namakoo-dev/basrun を"
                       " clone するか、環境変数 BASRUN でパスを指定して")
    return True, str(p)


def _check_demo_dir() -> tuple:
    d = HERE / "demo"
    ok = d.is_dir() and any(d.glob("*.xlsx"))
    detail = "" if ok else f"{d} に .xlsx サンプルが無い"
    return ok, detail


def doctor_checks(model: str = DEFAULT_MODEL) -> list:
    """(項目名, ok, 詳細/直し方) のリスト。判定ロジックだけを持ち、副作用(print)は
       cmd_doctor 側に置く（テストしやすくするため分離）。"""
    return [
        ("python 3.10+", *_check_python_version()),
        ("openpyxl", *_check_openpyxl()),
        (f"ollama 到達 ({OLLAMA})", *_check_ollama_reachable()),
        (f"モデル '{model}'", *_check_model_available(model)),
        ("LibreOffice", *_check_libreoffice()),
        ("basrun.py", *_check_basrun()),
        ("demo/", *_check_demo_dir()),
    ]


def format_doctor_report(results: list) -> tuple:
    """(表示テキスト, all_ok)。"""
    lines = []
    all_ok = True
    for name, ok, detail in results:
        mark = "✓" if ok else "×"
        if ok:
            line = f"{mark} {name}" + (f" ({detail})" if detail else "")
        else:
            all_ok = False
            line = f"{mark} {name}" + (f" — {detail}" if detail else "")
        lines.append(line)
    return "\n".join(lines), all_ok


def cmd_doctor(a: argparse.Namespace) -> int:
    text, all_ok = format_doctor_report(doctor_checks(a.model))
    print(text)
    return 0 if all_ok else 1


# ---------------------------------------------------------------------------
# 実行履歴（M1: 最小版・将来の需要センサー）
# ---------------------------------------------------------------------------

def build_history_entry(result: dict, book: Path, task: str, model: str, failure_kind: str,
                         error_detail: str | None = None) -> dict:
    """1 run の結果を history.jsonl の 1 行分の dict にする（純ロジック・テスト用に分離）。
       ★ M2a: error_detail は runtime_error 時の生エラー全文（端末には最終行だけ出す
       ため、詳細を追いたい時の唯一の入り口になる）。"""
    return {
        "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "book": str(book),
        "task": task,
        "model": model,
        "ok": bool(result.get("ok")),
        "attempts": result.get("attempts", 0),
        "failure_kind": failure_kind,
        "error_detail": error_detail,
        "changes": (result.get("changes") or [])[:3],
        "out": result.get("out"),
        # ★ M2b: DSL 経路(path="dsl")では命令言語の確認文(command)と事後条件の合否を残す。
        #   自由生成経路(path="freeform")では両方 None のまま（既存キーは不変）。
        "path": result.get("path", "freeform"),
        "command": result.get("command"),
        "postcondition": result.get("postcondition"),
    }


def append_history(entry: dict, path: Path | None = None) -> None:
    """history.jsonl に 1 行 append する。★ 失敗したら例外を投げる（run 本体を落とさ
       ないための try は呼び出し側(cmd_run)が持つ。ここでは書き込みロジックだけ）。"""
    p = path or HISTORY_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def read_history(path: Path | None = None, max_n: int = 10) -> list:
    """新しい順に最大 max_n 件を返す。壊れた行は読み飛ばす。"""
    p = path or HISTORY_FILE
    if not p.exists():
        return []
    entries = []
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return list(reversed(entries))[:max_n]


def format_history_table(entries: list) -> str:
    """人が読める表形式。履歴が無ければ「履歴はまだ無い」を返す。"""
    if not entries:
        return "履歴はまだ無い"
    header = f"{'日時':<20} {'結果':<4} {'試行':<4} {'モデル':<20} {'文書':<20} タスク"
    lines = [header]
    for e in entries:
        mark = "✓" if e.get("ok") else "×"
        ts = str(e.get("ts", ""))
        attempts = str(e.get("attempts", ""))
        model = str(e.get("model", ""))
        book = Path(str(e.get("book", ""))).name
        task = str(e.get("task", ""))
        line = f"{ts:<20} {mark:<4} {attempts:<4} {model:<20} {book:<20} {task}"
        kind = e.get("failure_kind")
        if not e.get("ok") and kind not in (None, "none"):
            line += f"  [{kind}]"
        lines.append(line)
    return "\n".join(lines)


def cmd_history(a: argparse.Namespace) -> int:
    entries = read_history(max_n=a.max)
    print(format_history_table(entries))
    return 0


# ---------------------------------------------------------------------------
# run コマンド本体
# ---------------------------------------------------------------------------

def _finish_run(a: argparse.Namespace, book: Path, result: dict, failure_kind: str,
                 error_detail: str | None = None) -> None:
    """--json 出力・成功時の注意書き・履歴の記録。cmd_run_freeform / cmd_run_dsl / cmd_run_plan
       の共通末尾。
       ★ DSL 経路(path="dsl")・複合計画経路(path="plan")は達成/総合判定の行を既に自分で
       出しているので、success_message() の『正しいかは差分を見て判断』（自由生成向けの
       注意書き）はここでは出さない。"""
    if a.json:
        print("\n" + json.dumps(result, ensure_ascii=False))
    if result.get("path") not in ("dsl", "plan"):
        msg = success_message(result)
        if msg:
            print("\n" + msg)
    try:
        detail = error_detail if error_detail is not None else (
            result.get("last_error_full") if failure_kind == "runtime_error" else None)
        append_history(build_history_entry(result, book, a.task, a.model, failure_kind,
                                            error_detail=detail))
    except Exception as e:
        print(f"WARN: 履歴の記録に失敗した: {e}", file=sys.stderr)


def cmd_run(a: argparse.Namespace) -> int:
    """run コマンドの入口。① 翻訳（計画）→
       - 計画が空/1段で CLARIFY → 質問して exit 3
       - 計画が空/1段で DSL 語彙 → ②〜⑥の決定論パイプライン(cmd_run_dsl)
       - 計画が空/1段でそれ以外(FREEFORM・翻訳失敗) → 現行の自由生成経路(cmd_run_freeform)
       - 計画が2段以上(複合依頼) → 段ごとに honest な項目別実行(cmd_run_plan)（M2c）
       ★ 後方互換: translate_task が "plan" で包まない旧形式（bare {"op":...}）を返した場合
       （テストの monkeypatch を含む）も、その dict をそのまま単一段として扱う。"""
    book = Path(a.book).resolve()
    if not book.exists():
        sys.exit(f"文書が無い: {book}")

    book_meta = build_book_meta(book)
    t0 = progress_start(f"⏳ 翻訳中 ({a.model})…")
    translation = translate_task(a.model, a.task, book_meta, temperature=0.1)
    progress_end(t0)

    plan = translation.get("plan") if isinstance(translation, dict) else None
    if not isinstance(plan, list) or not plan:
        if isinstance(translation, dict) and translation.get("op"):
            plan = [translation]
        else:
            plan = [{"op": "FREEFORM", "args": {}}]

    if len(plan) == 1:
        step = plan[0]
        op = step.get("op")
        if op == "CLARIFY":
            question = step.get("question") or "確認が必要です"
            print(f"？ {question}")
            return 3
        if op in OP_SCHEMA:
            return cmd_run_dsl(a, book, book_meta, op, step.get("args", {}))
        return cmd_run_freeform(a, book)

    return cmd_run_plan(a, book, book_meta, plan)


def _column_has_existing_values(book_path: Path, sheet_name: str, col_name: str) -> bool:
    """★ M2c: target(既存列指定)列に、見出し行を除いてどれか値が入っているか。
       上書き検知の明示用。読めない/列やシートが見つからない場合は False
       （保守的に『無い』扱い＝誤って警告しない）。"""
    try:
        wb = openpyxl.load_workbook(book_path, read_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return False
        ws = wb[sheet_name]
        idx = _col_index_by_header(ws, col_name)
        if idx is None:
            wb.close()
            return False
        last = _scan_last_row(ws)
        found = any(ws.cell(row=r, column=idx).value not in (None, "") for r in range(2, last + 1))
        wb.close()
        return found
    except Exception:
        return False


def _maybe_warn_target_overwrite(op: str, resolved: dict, book_meta: dict, book_path: Path) -> str | None:
    """★ M2c 項目2: COMPUTE_COLUMN の target(既存列指定)に既存値がある場合、
       上書きになる旨の1行を返す（無ければ None・確認行に明示するため）。"""
    if op != "COMPUTE_COLUMN" or not resolved.get("target"):
        return None
    sheets = book_meta.get("sheets") or []
    if not sheets:
        return None
    if _column_has_existing_values(book_path, sheets[0], resolved["target"]):
        return f"★ 対象列『{resolved['target']}』には既存値があります（上書きします）"
    return None


def cmd_run_dsl(a: argparse.Namespace, book: Path, book_meta: dict, op: str, raw_args: dict) -> int:
    """M2b の決定論パイプライン本体。②検証 → ③確認行 → ④codegen → ⑤適用 → ⑥事後条件。"""
    ok, resolved, inferred, err = verify_dsl_args(op, raw_args, book_meta)
    if not ok:
        print(f"？ {err}")
        return 3

    line = format_confirmation_line(op, resolved, inferred)
    print(f"■ ailine（DSL 経路）  model={a.model}  book={book.name}")
    print(line)
    warn_overwrite = _maybe_warn_target_overwrite(op, resolved, book_meta, book)
    if warn_overwrite:
        print(warn_overwrite)

    if a.ask:
        try:
            ans = input("続行しますか？ [y/N]: ").strip().lower()
        except EOFError:
            ans = ""
        if ans not in ("y", "yes"):
            print("× 中止した")
            return 1

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    out_book = book.with_name(book.stem + ".out" + book.suffix)
    apply_timeout = a.timeout if a.timeout else None   # 0 で無効化（旧挙動 = 無制限）
    helpers_dir = Path(a.helpers).resolve() if a.helpers else DEFAULT_HELPERS
    _helper_catalog, helper_files = load_helpers(helpers_dir)

    code = codegen_dsl(op, resolved, book_meta)
    (workdir / "dsl_attempt.bas").write_text(code, encoding="utf-8")
    print(f"\n─ 生成した .bas（決定論・LLM不使用）───────────────")
    print(code)
    print("──────────────────────────────────────────")

    result = {"ok": False, "attempts": 1, "task": a.task, "model": a.model,
              "path": "dsl", "command": line, "postcondition": None}

    if a.dry:
        print("（--dry: 適用しない。レビュー後に --dry を外して実行）")
        result["ok"] = True
        result["dry"] = True
        _finish_run(a, book, result, "none")
        return 0

    t0 = progress_start("⏳ 初回準備（文書の正規化）…")
    source_book = normalize_book(book, workdir, timeout=apply_timeout)
    progress_end(t0)
    before = snapshot(source_book)

    shutil.copy2(source_book, out_book)   # 原本は触らず、正規化済みコピーに適用
    t0 = progress_start("⏳ LibreOffice で適用中…")
    okrun, err_apply, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
    progress_end(t0)
    if not okrun:
        print(f"× 実行時エラー: {short_error_summary(err_apply)}（詳細は履歴に記録）。")
        result["last_error_full"] = err_apply
        _finish_run(a, book, result, "runtime_error", error_detail=err_apply)
        return 1

    after = snapshot(out_book)
    changed, lines = diff_snapshots(before, after)
    print("\n変更点:" if changed else "\n（文書に変化は検出されなかった）")
    for ln in lines:
        print(ln)
    advisories = build_advisories(a.task, before, after)
    for adv in advisories:
        print(adv)
    result["changes"] = lines
    result["advisories"] = advisories

    pcok, reason = run_postcondition(op, out_book, resolved, before_charts=before["charts"])
    result["postcondition"] = "pass" if pcok else "fail"
    if not pcok:
        print(f"\n× 適用されたが事後条件を満たさない: {reason}")
        result["out"] = str(out_book)
        _finish_run(a, book, result, "postcondition_fail")
        return 1

    print(f"\n✓ 達成を機械検証済み（操作:{OP_LABELS.get(op, op)}）: {reason}")
    result["ok"] = True

    if a.inplace:
        try:
            make_backup(book, keep=getattr(a, "keep_backups", DEFAULT_KEEP_BACKUPS))
        except Exception as e:
            print(f"× バックアップに失敗したため --inplace を中止した（原本は無変更）: {e}")
            print(f"適用先: {out_book.name}（--inplace は中止・原本 {book.name} は無変更）")
            result["out"] = str(out_book)
        else:
            shutil.move(out_book, book)
            print(f"\n適用先: {book.name}（--inplace で上書き）")
            print(f"復元: ailine restore {book.name}")
            result["out"] = str(book)
    else:
        print(f"\n適用先: {out_book.name}（原本 {book.name} は無変更）")
        result["out"] = str(out_book)

    _finish_run(a, book, result, "none")
    return 0


def cmd_run_freeform(a: argparse.Namespace, book: Path) -> int:
    """自由生成経路（従来の cmd_run 本体そのまま。M2a の助言つき）。
       ① 翻訳が CLARIFY にも DSL 語彙にも決まらなかった（FREEFORM・翻訳失敗）ときに使う。"""
    refs_dir = Path(a.refs).resolve() if a.refs else DEFAULT_REFS
    helpers_dir = Path(a.helpers).resolve() if a.helpers else DEFAULT_HELPERS
    helper_catalog, helper_files = load_helpers(helpers_dir)
    system = CONTRACT + load_refs(refs_dir) + helper_catalog
    desc = describe_book(book)
    user = f"{desc}\n\nタスク:\n{a.task}\n\n`Sub Run(oDoc As Object)` を1つだけ書け。コードのみ。"
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]

    print(f"■ ailine（自由生成経路）  model={a.model}  book={book.name}")
    print(f"■ 参照ライブラリ: {refs_dir}  ({len(list(refs_dir.glob('*.bas'))) if refs_dir.is_dir() else 0} 例)")
    print(f"■ ヘルパ: {helpers_dir}  ({len(helper_files)} 本を同梱・Call で呼ばせる)")

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    out_book = book.with_name(book.stem + ".out" + book.suffix)

    apply_timeout = a.timeout if a.timeout else None   # 0 で無効化（旧挙動 = 無制限）

    before = None
    source_book = book
    if not a.dry:
        t0 = progress_start("⏳ 初回準備（文書の正規化）…")
        source_book = normalize_book(book, workdir, timeout=apply_timeout)
        progress_end(t0)
        before = snapshot(source_book)

    result = {"ok": False, "attempts": 0, "task": a.task, "model": a.model,
              "path": "freeform", "command": None, "postcondition": None}
    failure_kind = "none"
    for attempt in range(a.repair + 1):
        result["attempts"] = attempt + 1
        t0 = progress_start(f"⏳ 生成中 ({a.model})…")
        raw = ollama_generate(a.model, msgs, temperature=a.temperature)
        progress_end(t0)
        code = extract_bas(raw)
        (workdir / f"attempt{attempt}.bas").write_text(code, encoding="utf-8")

        print(f"\n─ 試行 {attempt+1} ─ 生成した .bas ───────────────")
        print(code)
        print("──────────────────────────────────────────")

        if not valid_signature(code):
            print("× 署名が違う（Sub Run(oDoc As Object) が無い）。修復する。")
            failure_kind = "bad_signature"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": "署名が違う。`Sub Run(oDoc As Object)` を1つだけ。コードのみ。"}]
            continue

        # ★ M2a: bad_signature とは別の分類。署名はあるが本体が途中で切れているケース。
        if is_truncated_code(code):
            print("× 生成コードが不完全（途中で切断）。修復する。")
            failure_kind = "truncated"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "コードが途中で切れている（End Sub まで書き切れていない）。"
                      "最初から完全なコードを1つだけ書いて。コードのみ。"}]
            continue

        if a.dry:
            print("\n（--dry: 適用しない。レビュー後に --dry を外して実行）")
            result["ok"] = True
            result["dry"] = True
            failure_kind = "none"
            break

        shutil.copy2(source_book, out_book)   # 原本は触らず、正規化済みコピーに適用
        t0 = progress_start("⏳ LibreOffice で適用中…")
        ok, err, rawout = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not ok:
            # ★ M2a: obasync の生 Python トレースバックを端末にそのまま出さない。
            #   端末は最終行(例外名+メッセージ)だけ。修復ループへは従来どおり err の全文
            #   を渡す（モデルへの情報は減らさない）。全文は履歴 jsonl 側に残す。
            print(f"× 実行時エラー: {short_error_summary(err)}（詳細は履歴に記録）。修復する。")
            failure_kind = "runtime_error"
            result["last_error_full"] = err
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": f"実行時エラー: {err}\nこれを直して。コードのみ。"}]
            continue

        after = snapshot(out_book)
        changed, lines = diff_snapshots(before, after)
        if not changed:
            print("× no-op（実行は成功したが文書に変化が無い）。修復する。")
            failure_kind = "noop"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "実行は成功したが文書に一切変化が無かった（no-op）。"
                      "設定した API が効いていない可能性がある。別の正しい方法で書き直して。コードのみ。"}]
            continue

        print("\n✓ 適用され、文書が変化した。変更点:")
        for ln in lines:
            print(ln)
        advisories = build_advisories(a.task, before, after)
        for adv in advisories:
            print(adv)
        result["ok"] = True
        result["changes"] = lines
        result["advisories"] = advisories
        failure_kind = "none"
        if a.inplace:
            try:
                make_backup(book, keep=getattr(a, "keep_backups", DEFAULT_KEEP_BACKUPS))
            except Exception as e:
                # ★ M2a: バックアップ失敗時は --inplace を中止する（安全側・原本は無変更）。
                print(f"× バックアップに失敗したため --inplace を中止した（原本は無変更）: {e}")
                print(f"適用先: {out_book.name}（--inplace は中止・原本 {book.name} は無変更）")
                result["out"] = str(out_book)
            else:
                shutil.move(out_book, book)
                print(f"\n適用先: {book.name}（--inplace で上書き）")
                print(f"復元: ailine restore {book.name}")
                result["out"] = str(book)
        else:
            print(f"\n適用先: {out_book.name}（原本 {book.name} は無変更）")
            result["out"] = str(out_book)
        break
    else:
        print(f"\n× {a.repair+1} 回試みたが達成できなかった。")

    _finish_run(a, book, result, failure_kind)
    return 0 if result["ok"] else 1


# ---------------------------------------------------------------------------
# M2c: 複合依頼の計画実行と正直な範囲表示
#   翻訳(①)が返した plan(長さ2以上)を段ごとに実行する。DSL 語彙の段は②〜⑥の決定論
#   パイプライン、語彙外(OUT_OF_VOCAB/FREEFORM)の段は FREEFORM 経路（その段の依頼文だけ）。
#   ★ 黙落ゼロ: 計画に載った段は必ず項目別報告の1行になる。
#   ★ 総合判定は最弱の段に従う。「機械検証済み」の語は実際に機械検証が通った段にだけ付ける。
# ---------------------------------------------------------------------------

_ITEM_STATUS_MARK = {"ok": "✓", "warn": "⚠", "fail": "×"}

# col系 slot を持つ op → その slot 名（依存つき連鎖の新規列フォールバック対象）。
_COLUMN_ARG_KEYS = {
    "SORT": ("col",), "NUMBER_FORMAT": ("col",), "CHART": ("value_col",),
    "AGGREGATE": ("group_col", "value_col"),
}


def _apply_new_column_fallback(op: str, args: dict, headers: list, new_cols: list) -> dict:
    """★ M2c 依存つき連鎖（battery v2 #107 型）: 直前までの段が新規作成した列がちょうど
       1つあり、この段の列参照が現在の実列名のどれとも一致しない場合、その新規列を指して
       いるとみなして args を書き換える（候補が0か2つ以上なら何もしない＝保守的）。
       ★ 書き換えても最終的には verify_dsl_args が実在確認するので、誤った書き換えは
       通常どおりのエラーで止まる（無条件に信じ切らない）。"""
    if len(new_cols) != 1:
        return args
    only = new_cols[0]
    patched = dict(args)
    for k in _COLUMN_ARG_KEYS.get(op, ()):
        v = patched.get(k)
        if isinstance(v, str) and v not in headers and not re.fullmatch(r"\d+", v):
            patched[k] = only
    if op in ("BOLD", "FILL_COLOR", "CENTER_ALIGN"):
        t = patched.get("target", "")
        if isinstance(t, str) and t.startswith("col:"):
            name = t[4:]
            if name not in headers and not re.fullmatch(r"\d+", name):
                patched["target"] = f"col:{only}"
    return patched


def run_freeform_plan_step(a: argparse.Namespace, task_text: str, out_book: Path, workdir: Path,
                            refs_dir: Path, helpers_dir: Path, tag: str,
                            apply_timeout: float | None) -> tuple:
    """M2c: 複合計画の語彙外(OUT_OF_VOCAB/FREEFORM)段を FREEFORM 経路で実行する。
       cmd_run_freeform と同じ生成→適用→署名/切断/no-op チェックのループを、
       『その段の依頼文だけ』かつ『out_book の現在の状態』を起点に行う版。
       ★ cmd_run_freeform 本体は変えない（既存の回帰リスクを避けるため意図的に複製する）。
       戻り値: (ok, changes:list[str], advisories:list[str], failure_kind:str, detail:str|None)"""
    helper_catalog, helper_files = load_helpers(helpers_dir)
    system = CONTRACT + load_refs(refs_dir) + helper_catalog
    desc = describe_book(out_book)
    user = f"{desc}\n\nタスク:\n{task_text}\n\n`Sub Run(oDoc As Object)` を1つだけ書け。コードのみ。"
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]

    stepsource = workdir / f"{tag}_source{out_book.suffix}"
    shutil.copy2(out_book, stepsource)
    before = snapshot(stepsource)

    failure_kind = "none"
    for attempt in range(a.repair + 1):
        t0 = progress_start(f"⏳ 生成中（語彙外段・{a.model}）…")
        raw = ollama_generate(a.model, msgs, temperature=a.temperature)
        progress_end(t0)
        code = extract_bas(raw)
        (workdir / f"{tag}_attempt{attempt}.bas").write_text(code, encoding="utf-8")

        if not valid_signature(code):
            failure_kind = "bad_signature"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": "署名が違う。`Sub Run(oDoc As Object)` を1つだけ。コードのみ。"}]
            continue
        if is_truncated_code(code):
            failure_kind = "truncated"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "コードが途中で切れている（End Sub まで書き切れていない）。"
                      "最初から完全なコードを1つだけ書いて。コードのみ。"}]
            continue

        shutil.copy2(stepsource, out_book)
        t0 = progress_start("⏳ LibreOffice で適用中…")
        ok, err, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not ok:
            failure_kind = "runtime_error"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": f"実行時エラー: {err}\nこれを直して。コードのみ。"}]
            continue

        after = snapshot(out_book)
        changed, lines = diff_snapshots(before, after)
        if not changed:
            failure_kind = "noop"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "実行は成功したが文書に一切変化が無かった（no-op）。"
                      "設定した API が効いていない可能性がある。別の正しい方法で書き直して。コードのみ。"}]
            continue

        advisories = build_advisories(task_text, before, after)
        return True, lines, advisories, "none", None

    detail = {
        "bad_signature": "生成コードの署名が不正でした",
        "truncated": "生成コードが途中で切断されました",
        "runtime_error": "実行時エラーが解消しませんでした",
        "noop": "適用しても文書に変化がありませんでした",
    }.get(failure_kind, "原因不明で失敗しました")
    return False, [], [], failure_kind, detail


def format_plan_report(items: list) -> list:
    """複合計画の項目別報告を行のリストにする。items: [(idx, label, status, detail), ...]
       status は 'ok'/'warn'/'fail'。★ FREEFORM 段の成功は『機械検証済み』とは絶対に言わない
       （✓ 適用され文書が変化した級に留める＝warn 表示の固定文言で担保）。"""
    lines = []
    for idx, label, status, detail in items:
        mark = _ITEM_STATUS_MARK[status]
        if status == "ok":
            suffix = f"（{detail}）" if detail else ""
            lines.append(f"{idx}. {label} → {mark} 機械検証済み{suffix}")
        elif status == "warn":
            lines.append(f"{idx}. {label} → {mark} 語彙外のため自由生成で実行（確認してください）")
        else:
            lines.append(f"{idx}. {label} → {mark} 未対応: {detail}")
    return lines


def overall_verdict(items: list) -> tuple:
    """(判定文, 総合status)。★ 総合判定は最弱の段に従う:
       全段 ok → 「✓ すべて機械検証済み」/ fail 無しで warn を含む → 「⚠ 一部は確認が必要です」/
       fail を含む → 失敗。『達成を機械検証済み』の語は機械検証が実際に通った段にだけ付ける
       （ここでは全段が ok の時だけそう言う）。"""
    statuses = {it[2] for it in items}
    if "fail" in statuses:
        return "× 一部の操作が未対応/失敗のため、達成できませんでした", "fail"
    if "warn" in statuses:
        return "⚠ 一部は確認が必要です（語彙外の段は自由生成で実行・機械検証はしていません）", "warn"
    return "✓ すべて機械検証済み", "ok"


def cmd_run_plan(a: argparse.Namespace, book: Path, book_meta: dict, plan: list) -> int:
    """M2c: 複合依頼の計画実行本体。段ごとに②検証→③確認→④codegen→⑤適用→⑥事後条件
       （DSL 語彙の段）または FREEFORM（語彙外の段・その段の依頼文だけを渡す）を順に実行し、
       ★ 項目別の honest な報告を出す。総合判定は最弱の段に従う（cmd_run_plan 直上の
       overall_verdict）。
       ★ 依存つき連鎖: 各段の接地(verify_dsl_args)は直前までの段を実際に適用した後の
       out_book を読み直した列構成(current_meta)で行う。列名が一致しない場合は
       _apply_new_column_fallback が『直前段が作った新規列』への参照とみなして1回だけ
       書き換えを試みる。"""
    print(f"■ ailine（複合計画・{len(plan)} 段）  model={a.model}  book={book.name}")

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    out_book = book.with_name(book.stem + ".out" + book.suffix)
    apply_timeout = a.timeout if a.timeout else None
    helpers_dir = Path(a.helpers).resolve() if a.helpers else DEFAULT_HELPERS
    refs_dir = Path(a.refs).resolve() if a.refs else DEFAULT_REFS
    _helper_catalog, helper_files = load_helpers(helpers_dir)

    result = {"ok": False, "attempts": 1, "task": a.task, "model": a.model,
              "path": "plan", "command": None, "postcondition": None}

    if a.dry:
        print("\n（--dry プレビュー・語彙外の段は実行時に自由生成で対応します。未実行）")
        preview_items = []
        plan_json = []
        for i, step in enumerate(plan, 1):
            op = step.get("op")
            if op == "CLARIFY":
                q = step.get("question") or "確認が必要です"
                preview_items.append((i, q, "fail", "計画の途中で確認が必要なため対応できません"))
                plan_json.append({"op": "CLARIFY", "command": None, "status": "fail", "postcondition": None})
            elif op not in OP_SCHEMA:
                about = step.get("about") or "内容不明の依頼"
                preview_items.append((i, about, "warn", None))
                plan_json.append({"op": op, "command": about, "status": "warn", "postcondition": None})
            else:
                ok_v, resolved, inferred, err = verify_dsl_args(op, step.get("args", {}), book_meta)
                if ok_v:
                    label = format_confirmation_line(op, resolved, inferred)[len("解釈: "):]
                    preview_items.append((i, label, "ok", "未実行・プレビューのみ"))
                    plan_json.append({"op": op, "command": label, "status": "ok", "postcondition": None})
                else:
                    preview_items.append((i, f"操作:{OP_LABELS.get(op, op)}", "fail", err))
                    plan_json.append({"op": op, "command": None, "status": "fail", "postcondition": None})
        for ln in format_plan_report(preview_items):
            print(ln)
        print("\n（--dry: 適用しない。レビュー後に --dry を外して実行）")
        result["ok"] = True
        result["dry"] = True
        result["plan"] = plan_json
        _finish_run(a, book, result, "none")
        return 0

    t0 = progress_start("⏳ 初回準備（文書の正規化）…")
    source_book = normalize_book(book, workdir, timeout=apply_timeout)
    progress_end(t0)
    shutil.copy2(source_book, out_book)

    original_headers = {k: list(v) for k, v in book_meta["headers"].items()}
    first_sheet = book_meta["sheets"][0] if book_meta.get("sheets") else None
    before_all = snapshot(out_book)
    before_charts = before_all["charts"]

    current_meta = book_meta
    items: list = []         # (idx, label, status, detail)
    plan_json: list = []     # --json 用（既存キー不変・新規追加）

    for i, step in enumerate(plan, 1):
        op = step.get("op")

        if op == "CLARIFY":
            question = step.get("question") or "確認が必要です"
            items.append((i, question, "fail", "計画の途中で確認が必要なため対応できません"))
            plan_json.append({"op": "CLARIFY", "command": None, "status": "fail", "postcondition": None})
            continue

        if op not in OP_SCHEMA:
            about = step.get("about") or "内容不明の依頼"
            okf, changes, advisories, _fkind, detail = run_freeform_plan_step(
                a, about, out_book, workdir, refs_dir, helpers_dir, f"plan{i}", apply_timeout)
            if okf:
                items.append((i, about, "warn", None))
                for ln in changes:
                    print(f"  {ln}")
                for adv in advisories:
                    print(f"  {adv}")
            else:
                items.append((i, about, "fail", detail))
            plan_json.append({"op": op, "command": about,
                               "status": "ok" if okf else "fail", "postcondition": None})
            current_meta = build_book_meta(out_book)
            continue

        # 依存つき連鎖: 直前までの段の適用後の実列構成(current_meta)で接地する
        new_cols = []
        if first_sheet:
            new_cols = [c for c in current_meta["headers"].get(first_sheet, [])
                        if c not in original_headers.get(first_sheet, [])]
        raw_args = step.get("args", {})
        ok_v, resolved, inferred, err = verify_dsl_args(op, raw_args, current_meta)
        if not ok_v and new_cols and first_sheet:
            patched = _apply_new_column_fallback(
                op, raw_args, current_meta["headers"].get(first_sheet, []), new_cols)
            if patched != raw_args:
                ok_v2, resolved2, inferred2, err2 = verify_dsl_args(op, patched, current_meta)
                if ok_v2:
                    ok_v, resolved, inferred, err = ok_v2, resolved2, inferred2, err2

        if not ok_v:
            items.append((i, f"操作:{OP_LABELS.get(op, op)}", "fail", err))
            plan_json.append({"op": op, "command": None, "status": "fail", "postcondition": None})
            continue

        line = format_confirmation_line(op, resolved, inferred)
        label = line[len("解釈: "):]
        warn_overwrite = _maybe_warn_target_overwrite(op, resolved, current_meta, out_book)
        if warn_overwrite:
            print(f"  {i}段目: {warn_overwrite}")
        code = codegen_dsl(op, resolved, current_meta)
        (workdir / f"plan_step{i}.bas").write_text(code, encoding="utf-8")

        t0 = progress_start(f"⏳ {i}段目 LibreOffice で適用中…")
        okrun, err_apply, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not okrun:
            detail = f"実行時エラー: {short_error_summary(err_apply)}"
            items.append((i, label, "fail", detail))
            plan_json.append({"op": op, "command": line, "status": "fail", "postcondition": None})
            continue

        pcok, reason = run_postcondition(op, out_book, resolved, before_charts=before_charts)
        if not pcok:
            items.append((i, label, "fail", reason))
            plan_json.append({"op": op, "command": line, "status": "fail", "postcondition": "fail"})
            continue

        items.append((i, label, "ok", reason))
        plan_json.append({"op": op, "command": line, "status": "ok", "postcondition": "pass"})
        current_meta = build_book_meta(out_book)

    print()
    for ln in format_plan_report(items):
        print(ln)
    verdict_line, verdict = overall_verdict(items)
    print(f"\n{verdict_line}")

    after_all = snapshot(out_book)
    _changed, difflines = diff_snapshots(before_all, after_all)
    result["plan"] = plan_json
    result["items"] = [{"idx": idx, "label": label, "status": st, "detail": det}
                        for idx, label, st, det in items]
    result["changes"] = difflines

    if verdict == "fail":
        result["out"] = str(out_book)
        _finish_run(a, book, result, "plan_step_failed")
        return 1

    result["ok"] = True
    if a.inplace:
        try:
            make_backup(book, keep=getattr(a, "keep_backups", DEFAULT_KEEP_BACKUPS))
        except Exception as e:
            print(f"\n× バックアップに失敗したため --inplace を中止した（原本は無変更）: {e}")
            print(f"適用先: {out_book.name}（--inplace は中止・原本 {book.name} は無変更）")
            result["out"] = str(out_book)
        else:
            shutil.move(out_book, book)
            print(f"\n適用先: {book.name}（--inplace で上書き）")
            print(f"復元: ailine restore {book.name}")
            result["out"] = str(book)
    else:
        print(f"\n適用先: {out_book.name}（原本 {book.name} は無変更）")
        result["out"] = str(out_book)

    _finish_run(a, book, result, "none")
    return 0


def cmd_stop(a: argparse.Namespace) -> int:
    subprocess.run([sys.executable, str(basrun_path()), "stop"], encoding="utf-8", errors="replace")
    return 0


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="ailine", description="自然言語 → LibreOffice Basic → 適用 → 検証")
    sub = ap.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("run", help="タスクを生成・適用・検証する")
    r.add_argument("book", help="対象の文書 (.xlsx / .ods)")
    r.add_argument("task", help="やりたいことを自然言語で")
    r.add_argument("--model", default=DEFAULT_MODEL, help=f"ollama モデル (既定 {DEFAULT_MODEL})")
    r.add_argument("--refs", default=None, help="参照ライブラリのディレクトリ (既定 ./refs)")
    r.add_argument("--helpers", default=None, help="検証済みヘルパのディレクトリ (既定 ./helpers)")
    r.add_argument("--repair", type=int, default=2, help="修復の最大回数 (既定 2)")
    r.add_argument("--temperature", type=float, default=0.2)
    r.add_argument("--dry", action="store_true", help="生成して見せるだけ（適用しない）")
    r.add_argument("--inplace", action="store_true", help="原本を上書き（既定はコピー .out に適用）")
    r.add_argument("--json", action="store_true", help="結果を JSON でも出す")
    r.add_argument("--timeout", type=float, default=DEFAULT_APPLY_TIMEOUT,
                   help=f"basrun apply のタイムアウト秒 (既定 {DEFAULT_APPLY_TIMEOUT:.0f}、"
                        "0 で無効化=旧挙動の無制限)")
    r.add_argument("--ask", action="store_true",
                   help="DSL 経路の確認行の後に y/n で対話する（既定は表示して続行）")
    r.add_argument("--keep-backups", dest="keep_backups", type=int, default=DEFAULT_KEEP_BACKUPS,
                   help=f"--inplace のバックアップを book ごとに何世代残すか (既定 {DEFAULT_KEEP_BACKUPS}、"
                        "負数で無制限)")
    r.set_defaults(func=cmd_run)

    s = sub.add_parser("stop", help="起動した LibreOffice を落とす")
    s.set_defaults(func=cmd_stop)

    d = sub.add_parser("doctor", help="セットアップを診断する")
    d.add_argument("--model", default=DEFAULT_MODEL, help=f"確認するモデル (既定 {DEFAULT_MODEL})")
    d.set_defaults(func=cmd_doctor)

    h = sub.add_parser("history", help="実行履歴を表示する")
    h.add_argument("--max", type=int, default=10, help="表示件数（既定 10、新しい順）")
    h.set_defaults(func=cmd_history)

    rs = sub.add_parser("restore", help="--inplace のバックアップから復元する")
    rs.add_argument("book", help="対象の文書 (.xlsx / .ods)")
    rs.add_argument("--list", action="store_true", help="バックアップ一覧を表示するだけ（復元しない）")
    rs.set_defaults(func=cmd_restore)
    return ap


def main(argv=None) -> int:
    a = build_parser().parse_args(argv)
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())
