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

# ★ A': 用語集（税率等の「現場の取り決め値」）。グローバルのみ・ブック別上書きは作らない
#   （YAGNI・受信ファイル注入の予防。サイドカー自動読みは絶対にしない）。
VOCAB_FILE = HISTORY_DIR / "vocab.json"
DEFAULT_VOCAB_MAX_ENTRIES = 200     # 個人利用の上限。無制限にしない
DEFAULT_VOCAB_MAX_TERM_LEN = 40     # 語（キー）の最大長


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
- ★ 新しいシートを作らず、既存シートの中の最小の変更で達成することを最優先する
  （依頼が明示的にシート新設・ピボットを求めていない限り）。
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
        "★ headerRow 引数は見出し行（0起点）。見出しが物理1行目ならほぼ常に 0。\n"
        "例: 列0〜4の表で金額が列1なら、金額で降順に並べ替え"
        " → `Call SortByColumn(oDoc, 0, 4, 1, False)`（第2引数 lastCol=表の最終列）\n"
        "例: 金額(列1)の棒グラフ（項目名は先頭列に自動）→ `Call InsertBarChart(oDoc, 0, 1)`\n"
        "例: A1とB1を結合 → `Call MergeCells(oDoc, 0, 0, 1, 0)`\n"
        "例: 先頭データ行(2行目)の前に1行挿入 → `Call InsertRows(oDoc, 1, 1)`\n"
        "例: 表に罫線を引く → `Call DrawTableBorders(oDoc)`\n"
        "例: 各列の幅を内容に合わせる → `Call AutoFitColumns(oDoc)`\n"
        "例: C列(列2)に、商品名(列0)をキーに『単価表』から値を引く（VLOOKUP相当）"
        " → `Call VLookupFromTable(oDoc, 0, 0, 2, \"単価表\")`（参照表は 列0=キー・列1=値）\n"
        "例: 『ピボット』で部門(列0)ごとに金額(列1)を集計（本物の DataPilot・Excel で操作可）"
        " → `Call PivotSum(oDoc, 0, 1)`\n"
        "例: 『集計表／まとめ』を作る＝部門(列0)ごとの金額(列1)を見栄えのする普通の表に"
        "（罫線・カンマ・太字つき。★『ピボット』と明示されない集計は基本こちら）"
        " → `Call SummaryTable(oDoc, 0, 0, 1)`\n"
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


def book_columns(path: Path, header_rows: dict | None = None) -> dict:
    """全シートの見出し行を {シート名: [列名,...]} で返す。
       ★ describe_book は1枚目だけの人間可読版。こちらは M2b の翻訳・検証が使う
       機械可読の接地情報（列は最初の空欄で打ち切る＝連続した見出しだけを列とみなす）。
       ★ W3: header_rows({シート名: 見出し行(1起点)}) を渡すと、そのシートだけは
       物理1行目でなく指定行を見出しとして読む（StructDump の見出し検出結果を反映する）。
       未指定のシートは従来どおり1行目（後方互換・引数省略時は完全に旧挙動）。"""
    header_rows = header_rows or {}
    wb = openpyxl.load_workbook(path, read_only=True)
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        hr = header_rows.get(name, 1)
        ncol = min(ws.max_column or 0, MAX_COLS)
        headers = []
        for c in range(1, ncol + 1):
            v = ws.cell(row=hr, column=c).value
            if v in (None, "") and hr > 1:
                # ★ W3: 多段見出し対策。子見出し行(hr>1)で空欄の列は、真上の行を遡って
                #   最初に見つかる非空値を引き継ぐ（D検体: 商品名は1行目だけにあり
                #   2行目(子見出し)の同じ列は空。無いと打ち切り誤検知で列挙が0件になる）。
                for up in range(hr - 1, 0, -1):
                    uv = ws.cell(row=up, column=c).value
                    if uv not in (None, ""):
                        v = uv
                        break
            if v in (None, ""):
                break
            headers.append(str(v))
        out[name] = headers
    wb.close()
    return out


def build_book_meta(path: Path, header_rows: dict | None = None) -> dict:
    """{"sheets": [...], "headers": {シート名: [列名,...]}, "header_rows": {シート名: 行(1起点)}}。
       M2b 翻訳・検証の接地情報。★ W3: header_rows 省略時は全シート1行目（旧挙動と同一）。"""
    header_rows = dict(header_rows or {})
    headers = book_columns(path, header_rows)
    resolved_header_rows = {name: header_rows.get(name, 1) for name in headers}
    return {"sheets": list(headers.keys()), "headers": headers, "header_rows": resolved_header_rows}


# ---------------------------------------------------------------------------
# W3 Part1/2: StructDump（LibreOffice の目で構造を読む）+ 見出し行推定
#
# ★ 正規化パス(normalize_book)の LO 往復に同乗させる（二役）。normalize_book が実行する
#   マクロを「何もしない空マクロ」から「何もしないが構造をテキストへ書き出すマクロ」に
#   差し替えるだけで、追加の LO 起動は発生しない。
# ★ LibreOffice の Cursor.gotoEndOfUsedArea・図形/グラフ/DataPilot 個数は LO でしか
#   正確に取れない（openpyxl の ws.max_row は書式だけ残った幽霊セルで過大評価しうる）。
#   一方、行ごとの書式的特徴(太字数等)・結合範囲は、LO が保存し終えた同じファイルを
#   openpyxl で読むだけで求まる（もう一度 LO を起動しない）。
# ---------------------------------------------------------------------------

STRUCT_HEADER_SCAN_ROWS = 20   # 見出し検出に使う先頭行数（多段見出し・タイトル行を含めても十分な余裕）
STRUCTDUMP_FILENAME = "structdump.txt"   # Basic → Python の受け渡し用（生のタブ区切りテキスト）


def _structdump_macro(out_path: Path) -> str:
    """StructDump 用の Basic マクロ。文書には一切手を触れない（正規化パスの no-op 性質を
       保つ）。シート毎に「実使用範囲(Cursor.gotoEndOfUsedArea)・図形/グラフ/DataPilot 個数」
       だけをタブ区切りテキストで書き出す（Basic に JSON エンコードは無いため、
       機械可読 dict への組み立ては Python 側 build_struct_dump() が担う）。"""
    out_str = str(out_path).replace('"', '""')
    return f'''Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim iFile As Integer, i As Integer, n As Integer
    Dim oSheet As Object, oCursor As Object, oAddr As Object
    iFile = FreeFile
    Open "{out_str}" For Output As #iFile
    n = oDoc.Sheets.Count
    For i = 0 To n - 1
        oSheet = oDoc.Sheets.getByIndex(i)
        oCursor = oSheet.createCursor()
        oCursor.gotoEndOfUsedArea(True)
        oAddr = oCursor.RangeAddress
        Print #iFile, "SHEET" & Chr(9) & oSheet.Name & Chr(9) & oAddr.StartColumn & Chr(9) _
            & oAddr.StartRow & Chr(9) & oAddr.EndColumn & Chr(9) & oAddr.EndRow & Chr(9) _
            & oSheet.DrawPage.Count & Chr(9) & oSheet.Charts.Count & Chr(9) & oSheet.DataPilotTables.Count
    Next i
    Close #iFile
End Sub
'''


def parse_structdump_raw(text: str) -> dict:
    """StructDump の生テキスト（"SHEET"行群・タブ区切り）を
       {シート名: {"used_range": {...0起点...}, "shapes":n, "charts":n, "datapilots":n}} にパースする。
       壊れた/欠けた行は無視する（安全側・落とさない）。"""
    sheets: dict = {}
    for line in text.splitlines():
        parts = line.split("\t")
        if len(parts) != 9 or parts[0] != "SHEET":
            continue
        try:
            name = parts[1]
            sc, sr, ec, er = int(parts[2]), int(parts[3]), int(parts[4]), int(parts[5])
            shapes, charts, pivots = int(parts[6]), int(parts[7]), int(parts[8])
        except ValueError:
            continue
        sheets[name] = {
            "used_range": {"start_col": sc, "start_row": sr, "end_col": ec, "end_row": er},
            "shapes": shapes, "charts": charts, "datapilots": pivots,
        }
    return sheets


def _row_char_stats(ws, start_row: int, end_row: int, start_col: int, end_col: int) -> dict:
    """StructDump の行ごとの書式的特徴（見出し検出のヒューリスティクスに足る最小限）。
       {行(1起点): {"nonempty": 非空セル数, "str": うち文字列セル数, "bold": 太字セル数}}。"""
    stats = {}
    for r in range(start_row, end_row + 1):
        nonempty = 0
        strcnt = 0
        boldcnt = 0
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v in (None, ""):
                continue
            nonempty += 1
            if isinstance(v, str):
                strcnt += 1
            if cell.font and cell.font.bold:
                boldcnt += 1
        stats[r] = {"nonempty": nonempty, "str": strcnt, "bold": boldcnt}
    return stats


def build_struct_dump(normalized_book: Path, workdir: Path) -> dict:
    """StructDump 本体。normalize_book() が同じ LO 往復で書き出した生ダンプ
       (workdir/STRUCTDUMP_FILENAME) をパースし、正規化済みブックを openpyxl で読んで
       行の書式的特徴・結合範囲を足した機械可読 dict にする（もう一度 LO を起動しない）。
       生ダンプが無い/壊れている場合はそのシートだけ openpyxl の推定で代用する
       （テストでの normalize_book 差し替え等・呼び出し側は既定 header_row=1 に退避できる）。
       戻り値: {"sheets": {シート名: {"used_range":{1起点}, "merges":[...], "shapes":n,
                "charts":n, "datapilots":n, "rows": {行(1起点): {...}}}}}"""
    dump_path = workdir / STRUCTDUMP_FILENAME
    raw_sheets: dict = {}
    if dump_path.exists():
        try:
            raw_sheets = parse_structdump_raw(dump_path.read_text(encoding="utf-8"))
        except Exception:
            raw_sheets = {}
    wb = openpyxl.load_workbook(normalized_book)
    out: dict = {"sheets": {}}
    for name in wb.sheetnames:
        ws = wb[name]
        raw = raw_sheets.get(name)
        if raw is not None:
            ur = raw["used_range"]
            start_row, end_row = ur["start_row"] + 1, ur["end_row"] + 1     # 0起点→1起点
            start_col, end_col = ur["start_col"] + 1, ur["end_col"] + 1
        else:
            start_row, start_col = 1, 1
            end_row = min(ws.max_row or 1, MAX_ROWS)
            end_col = min(ws.max_column or 1, MAX_COLS)
        scan_end = min(end_row, start_row + STRUCT_HEADER_SCAN_ROWS - 1)
        rows = _row_char_stats(ws, start_row, scan_end, start_col, min(end_col, MAX_COLS))
        out["sheets"][name] = {
            "used_range": {"start_row": start_row, "end_row": end_row,
                            "start_col": start_col, "end_col": end_col},
            "merges": sorted(str(r) for r in ws.merged_cells.ranges),
            "shapes": raw["shapes"] if raw else 0,
            "charts": raw["charts"] if raw else 0,
            "datapilots": raw["datapilots"] if raw else 0,
            "rows": rows,
        }
    wb.close()
    return out


def detect_header_row(sheet_struct: dict) -> tuple:
    """(見出し行(1起点) or None, confident: bool)。
       ヒューリスティクス: 「上に結合セルやタイトル行があっても、列方向に複数(2以上)の
       非空文字列セルが並び、その行自体は非空セルが全部文字列（=見出しらしい）で、
       直下の行に文字列でない非空セル(数値等)が混ざる（=データが始まる）」行を強い候補とする。
       ★ D検体（2段見出し）対策: 親見出し行(結合)も str_count>=2 を満たしうるが、その直下も
       まだ見出し（型の混在なし）なので候補から外れる。子見出し行の直下でようやくデータの
       型混在が起き、そこで初めて候補になる（追加ルール不要でこの一般則から自然に解ける）。
       ★ フォールバック: 型混在チェックで候補0件（例: 全列が文字列のみの表）でも、
       「非空文字列セル2以上・行自体は純文字列」を満たす行が候補全体でちょうど1つなら
       それを確信とする（曖昧さが無いケース）。
       強い候補が0個か2個以上（=曖昧）なら (None, False) を返す＝呼び出し側は推測せず CLARIFY。"""
    rows = sheet_struct.get("rows", {})
    if not rows:
        return None, False
    sorted_rows = sorted(rows.keys())
    pure_str_rows = [r for r in sorted_rows
                      if rows[r]["str"] >= 2 and rows[r]["nonempty"] == rows[r]["str"]]

    def _mixture_below(r: int) -> bool:
        nxt = rows.get(r + 1)
        return nxt is not None and nxt["nonempty"] > nxt["str"]

    with_mixture = [r for r in pure_str_rows if _mixture_below(r)]
    if len(with_mixture) == 1:
        return with_mixture[0], True
    if not with_mixture and len(pure_str_rows) == 1:
        return pure_str_rows[0], True
    return None, False


CLARIFY_HEADER_ROW_QUESTION = "見出しは何行目ですか？（1 行目/3 行目 のように答えて）"


def resolve_header_rows(struct_dump: dict, sheets: list) -> tuple:
    """全シートの見出し行(1起点)を決める。(header_rows: {シート名: 行}, clarify_question|None)。
       ★ 1枚目シートだけ StructDump のヒューリスティクスで推定する（DSL 操作は1枚目シートに
       限定されているため）。他シート（LOOKUP_FILL の参照表等）は物理1行目を既定にする。
       StructDump が無い（テストでの normalize_book 差し替え等）場合は全シート1行目のまま
       （旧挙動と同一・CLARIFY は出さない）。自信が持てない場合だけ1枚目シートについて
       CLARIFY 質問を返す（推測で進まない）。"""
    header_rows = {s: 1 for s in sheets}
    if not sheets:
        return header_rows, None
    sd_sheets = (struct_dump or {}).get("sheets", {})
    first = sheets[0]
    info = sd_sheets.get(first)
    if info is None:
        return header_rows, None
    row, confident = detect_header_row(info)
    if confident:
        header_rows[first] = row
        return header_rows, None
    return header_rows, CLARIFY_HEADER_ROW_QUESTION


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
    # ★ 止血3: truncated は「この snapshot が MAX_ROWS/MAX_COLS で切り詰められたか」を
    #   保持する。diff_snapshots 後の表示で「先頭1000行しか見せていない」ことを正直に
    #   注記するために使う（bench/realworld/BASELINE.md の B 検体所見の根治）。
    snap = {"sheets": list(wb.sheetnames), "charts": _charts_count(path),
            "cells": {}, "merges": {}, "colw": {}, "rowh": {}, "truncated": False}
    for name in wb.sheetnames:
        ws = wb[name]
        true_nrow = ws.max_row or 0
        true_ncol = ws.max_column or 0
        if true_nrow > MAX_ROWS or true_ncol > MAX_COLS:
            snap["truncated"] = True
        nrow = min(true_nrow, MAX_ROWS)
        ncol = min(true_ncol, MAX_COLS)
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


def _truncation_notice(before: dict, after: dict, exhaustive_postcondition: bool) -> str | None:
    """★ 止血3: before/after どちらかの snapshot が MAX_ROWS/MAX_COLS で切り詰められて
       いたら、無言で切らず正直な1行を返す（bench/realworld/BASELINE.md の B 検体所見）。
       ★ 経路で文言を出し分ける:
       - exhaustive_postcondition=True（DSL経路・cmd_run_dsl / cmd_run_plan の DSL 段）:
         事後条件チェッカーは openpyxl で out ファイルを直接開き _scan_last_row で全行を
         走査する（snapshot() の MAX_ROWS とは無関係）。表示だけが切り詰められている。
       - exhaustive_postcondition=False（FREEFORM経路）: no-op ガード・advisories も
         snapshot() 頼みなので、検証自体も先頭1000行までしか見ていない。
       切り詰めが無ければ None。"""
    if not (before.get("truncated") or after.get("truncated")):
        return None
    if exhaustive_postcondition:
        return f"（表示は先頭 {MAX_ROWS} 行の変化のみ。検証・適用は全行に対して実施）"
    return f"（表示は先頭 {MAX_ROWS} 行の変化のみ。検証も先頭 {MAX_ROWS} 行のみ）"


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
       ★ M2c: 判定対象は『値変更』の部分集合だけ（書式のみの変更は無視・保守性は部分集合内で維持）。
       ★ W6: 実行中に新規作成されたシート（before["sheets"] に無い）のセルはここでは無視する
       （new_sheet_advisories が別途担当）。以前は『使用範囲が不明なシートが1つでも混ざると
       関数全体が判定を保留する』実装だったため、AGGREGATE/CHART/PivotSum のように新規シート
       を作る操作が絡むたび、他シートの本当のゴーストデータ検出まで丸ごと素通りしていた
       （監査実測。旧シート範囲が不明＝原本に無い＝新規シート、の場合に限って除外することで
       既存シートに対する検出力は変えずに直す）。"""
    changed = _value_changed_cells(before, after)
    if not changed:
        return None
    outside = []
    for sheet, r, c in changed:
        if sheet not in before["sheets"]:
            continue   # ★ W6: 新規作成されたシート＝ここでの判定対象外
        rect = _used_range(before, sheet)
        if rect is None:
            return None  # このシートの原本データ範囲が不明 → 判定を保留
        min_r, max_r, min_c, max_c = rect
        if min_r <= r <= max_r and min_c <= c <= max_c:
            return None  # 1つでも範囲内 → 発火しない
        outside.append((r, c))
    if not outside:
        return None   # 変更が全部、新規シートのセルだけだった
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


def _new_sheets(before: dict, after: dict) -> list:
    """before に無く after に有るシート名（実行中に新規作成されたシート）。順序維持。"""
    return [s for s in after["sheets"] if s not in before["sheets"]]


def _sheet_only_snapshot(snap: dict, sheet: str) -> dict:
    """snap のうち sheet に属する情報だけを持つ縮小 snapshot。
       ★ detect_ghost_data/detect_uniform_fill/count_reconciliation は snapshot() が返す
       dict の形（cells/merges/colw/rowh/sheets）をそのまま前提にした既存ロジックなので、
       新規シート専用に書き直さず、この縮小 snapshot 越しに『再利用』する（new_sheet_advisories
       が使う）。"""
    prefix = sheet + "!"
    return {
        "sheets": [sheet],
        "charts": 0,
        "cells": {k: v for k, v in snap["cells"].items() if k.startswith(prefix)},
        "merges": {sheet: snap.get("merges", {}).get(sheet, [])},
        "colw": {sheet: snap.get("colw", {}).get(sheet, {})},
        "rowh": {sheet: snap.get("rowh", {}).get(sheet, {})},
        "truncated": snap.get("truncated", False),
    }


def new_sheet_advisories(before: dict, after: dict) -> list:
    """★ W6 項目2: 監査所見「新規『集計』シートの全 0 埋めが★素通り」の根治。
       detect_ghost_data / detect_uniform_fill / count_reconciliation は『変更セル全部が
       該当した時だけ発火する』設計のため、新規シートの異常が他シートの正常な変更と混ざると
       『全部該当』が崩れ、丸ごと素通りしていた（例: COMPUTE_COLUMN が本シートの小計に書いた
       正しい値と、AGGREGATE が作った『集計』シートの壊れた全0埋めが同じ diff に同居する）。
       新規シートごとに単独の縮小 before/after（空 before・そのシートだけの after）を作って
       同じ3関数へ通すことで、他シートの変更に影響されず判定する。
       ★ detect_ghost_data は『原本の使用範囲外か』を測る関数だが、新規シートに原本の
       使用範囲は存在しない（何もかもが『新規』であって『範囲外』ではない）。そのため
       常に None を返す（適用はするが構造的に無言＝その関数の既存の誠実さをそのまま踏襲）。"""
    lines = []
    for sheet in _new_sheets(before, after):
        empty_before = {"sheets": [sheet], "charts": 0, "cells": {}, "merges": {sheet: []},
                         "colw": {sheet: {}}, "rowh": {sheet: {}}, "truncated": False}
        sheet_after = _sheet_only_snapshot(after, sheet)
        for fn in (detect_ghost_data, detect_uniform_fill):
            msg = fn(empty_before, sheet_after)
            if msg:
                lines.append(msg.replace("★ 疑わしい: ", f"★ 疑わしい: 新規シート『{sheet}』の", 1))
        recon = count_reconciliation(empty_before, sheet_after)
        if recon:
            lines.append(f"新規シート『{sheet}』 {recon}")
    return lines


_NEW_SHEET_MENTION_RE = re.compile(r"シート|ピボット|別に")


def unrequested_new_sheet_advisory(task: str, before: dict, after: dict) -> list:
    """★ W6 項目3（機械側）: 依頼文にシート新設の明示的な言及（『シート』『ピボット』
       『別に』のいずれか）が無いのに新規シートが作られたら申告する。
       ★ 保守的: 言及があれば（AGGREGATE/CHART/PivotSum 等が意図どおり新設したと見なし）沈黙。
       プロンプト側の抑制（CONTRACT の追記）はあくまで誘導であって保証にならないため、
       この機械申告が最終防衛線（feedback_intent_vs_guarantee: 指示は意図、保証は機械）。"""
    new_sheets = _new_sheets(before, after)
    if not new_sheets or _NEW_SHEET_MENTION_RE.search(task):
        return []
    return [f"★ 依頼にない新しいシートが作成されました（{s}）" for s in new_sheets]


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
       ①幽霊データ ②一様埋め ③件数の突き合わせ ④依頼文言との重なり
       ⑤新規シートの中身（★ W6） ⑥依頼にないシート新設の申告（★ W6）。"""
    lines = []
    for fn in (detect_ghost_data, detect_uniform_fill):
        msg = fn(before, after)
        if msg:
            lines.append(msg)
    recon = count_reconciliation(before, after)
    if recon:
        lines.append(recon)
    lines.extend(new_sheet_advisories(before, after))
    lines.extend(unrequested_new_sheet_advisory(task, before, after))
    mentions = extract_task_mentions(task, before["sheets"])
    lines.extend(mention_overlap_advisory(mentions, before, after))
    return lines


# ---------------------------------------------------------------------------
# ★ A': 用語集（vocab）— 税率等の「現場の取り決め値」を LLM でなく辞書に持たせる。
#   ~/.ailine/vocab.json（グローバルのみ）に {"語": 値} の平坦な dict で持つ。
#   設定ファイル経由で唯一「自由入力」がコード生成に届く経路なので、他のどの入力より
#   厳しく検疫する（float() 関門・制御文字禁止・件数/長さ上限。壊れたファイルは
#   クラッシュせず空辞書として扱う）。
# ---------------------------------------------------------------------------

def _sanitize_vocab_term(term) -> str | None:
    """語（キー）が登録可能な形か。空・制御文字・長すぎるものは None（拒否）。"""
    s = str(term).strip()
    if not s or len(s) > DEFAULT_VOCAB_MAX_TERM_LEN:
        return None
    if re.search(r"[\x00-\x1f\x7f]", s):   # 改行等の制御文字禁止（codegen へ渡る経路の防御）
        return None
    return s


def load_vocab(path: Path | None = None) -> dict:
    """~/.ailine/vocab.json を読む。無い/壊れている/形が違う場合は空の辞書を返す
       （★ クラッシュしない・battery や pytest からは path 明示で差し替えて再現性を保つ）。
       エントリごとに語をサニタイズし、値は float() を通ったものだけを採用する。
       件数が上限を超えた分は読み捨てる（先着順・ファイルの並び順に依存）。"""
    p = path or VOCAB_FILE
    if not p.is_file():
        return {}
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(raw, dict):
        return {}
    vocab: dict = {}
    for term, value in raw.items():
        if len(vocab) >= DEFAULT_VOCAB_MAX_ENTRIES:
            break
        clean_term = _sanitize_vocab_term(term)
        if clean_term is None:
            continue
        try:
            vocab[clean_term] = float(value)
        except (TypeError, ValueError):
            continue
    return vocab


def save_vocab(vocab: dict, path: Path | None = None) -> None:
    """vocab を ~/.ailine/vocab.json に上書き保存する。"""
    p = path or VOCAB_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(vocab, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def vocab_add(term: str, value, path: Path | None = None) -> tuple:
    """(ok, message)。term/value を検証してから登録・保存する。
       未登録の新規語で件数が上限に達している場合は拒否する（既存語の値更新は上限に関係なく可）。"""
    clean_term = _sanitize_vocab_term(term)
    if clean_term is None:
        return False, f"語『{term}』は登録できません（空/制御文字/{DEFAULT_VOCAB_MAX_TERM_LEN}文字超）"
    try:
        fval = float(value)
    except (TypeError, ValueError):
        return False, f"値『{value}』が数値ではありません"
    vocab = load_vocab(path)
    if clean_term not in vocab and len(vocab) >= DEFAULT_VOCAB_MAX_ENTRIES:
        return False, f"用語集が上限（{DEFAULT_VOCAB_MAX_ENTRIES}件）に達しています"
    vocab[clean_term] = fval
    save_vocab(vocab, path)
    return True, f"登録: {clean_term} = {fval:g}"


# --- ★ A': APPEND_TOTAL の倍率(factor)を LLM から切り離し、機械が確定する ------------
#   ①依頼文の明示率を regex で抽出 ②無ければ用語集を引く ③どちらも無く label が税/込を
#   含むなら CLARIFY。LLM が factor を返しても、ここが常に勝つ（verify_dsl_args 側で
#   食い違いを WARN として記録する）。

_RATE_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[%％]")
_RATE_BAI_RE = re.compile(r"(\d+(?:\.\d+)?)\s*倍")
_RATE_KEYWORD_RE = re.compile(r"税|倍率")
_RATE_BARE_NUM_RE = re.compile(r"(\d+(?:\.\d+)?)")


def extract_rate_factor(text: str) -> tuple:
    """依頼文から明示の倍率を抽出する。戻り値は (factor, 出典スニペット) か (None, None)。
       ①「10%」「8 ％」型 → 1+n/100 ②「1.1倍」型 → n そのまま ③「税」「倍率」という語の
       前後8文字だけにある裸の小数（例:「税率0.1」）→ 1未満なら 1+n・1以上ならそのまま
       （無関係な数値の誤爆を避けるため、③だけは税/倍率の語の近傍に絞る）。
       複数の異なる値が見つかった場合は断定しない（None, None・CLARIFY に委ねる）。"""
    if not text:
        return None, None
    candidates: dict = {}   # factor -> 出典スニペット（最初に見つかったもの）
    for m in _RATE_PCT_RE.finditer(text):
        f = round(1 + float(m.group(1)) / 100, 6)
        candidates.setdefault(f, m.group(0))
    for m in _RATE_BAI_RE.finditer(text):
        f = round(float(m.group(1)), 6)
        candidates.setdefault(f, m.group(0))
    if not candidates:
        for km in _RATE_KEYWORD_RE.finditer(text):
            window = text[max(0, km.start() - 8): km.end() + 8]
            nm = _RATE_BARE_NUM_RE.search(window)
            if nm:
                n = float(nm.group(1))
                f = round((1 + n) if n < 1 else n, 6)
                candidates.setdefault(f, nm.group(0))
    if len(candidates) == 1:
        f, snippet = next(iter(candidates.items()))
        return f, snippet
    return None, None


def lookup_vocab_factor(text: str, vocab: dict) -> tuple:
    """依頼文に用語集の語が部分一致で含まれるかを見る。戻り値は (factor, 用語) か
       (None, None)。複数の異なる語（異なる値）がヒットした場合は断定しない。"""
    if not text or not vocab:
        return None, None
    hits: dict = {}
    for term, value in vocab.items():
        if term and term in text:
            hits.setdefault(value, term)
    if len(hits) == 1:
        value, term = next(iter(hits.items()))
        return value, term
    return None, None


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
    "CENTER_ALIGN": "中央揃え", "APPEND_TOTAL": "合計追加",
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
    # ★ W6: label/factor は既定値がある任意項目（無指定でも FREEFORM に退避させない）。
    "APPEND_TOTAL": ("col",),
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
CENTER_ALIGN: 中央揃え。args: target("all" か "col:列名")
APPEND_TOTAL: 列の合計を表の最終行の下に追加する（税込み合計等）。args: col(合計する列名),
  label(省略可・既定"合計"。表示ラベル。「税込み合計」等、依頼の言い方をそのまま入れる)
  ★ 倍率(税率等)は入れない。数値化はここでは行わない（機械が別途確定する）"""

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
    # ★ W6: APPEND_TOTAL 語彙昇格（監査3回連続失敗の実測による）。
    #   ①明示的な倍率(消費税等)つき ②単純な合計、の両方を教える。
    #   ★ A': 倍率の数値化(1.1等)は LLM に求めない（factor は machine-determined。
    #   verify_dsl_args の extract_rate_factor/lookup_vocab_factor が依頼文から確定する）。
    ('対象ブックの構成: {"Sheet": ["品目", "数量", "単価", "小計"]}\n'
     '依頼: 「税込み合計を一番下に出して（消費税10%）」',
     '{"plan": [{"op": "APPEND_TOTAL", "args": '
     '{"col": "小計", "label": "税込み合計"}}]}'),
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「金額の合計を最後に」',
     '{"plan": [{"op": "APPEND_TOTAL", "args": {"col": "金額"}}]}'),
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


def verify_dsl_args(op: str, args: dict, book_meta: dict, task: str = "", vocab: dict | None = None) -> tuple:
    """② 検証。(ok, resolved_args, inferred_keys, error_message)。
       args のシート/列名が実在するかを機械照合し、実在名に解決する。実在しなければ
       CLARIFY 相当のエラーメッセージを返す（呼び出し側が確認質問として表示する）。
       ★ A': task/vocab は APPEND_TOTAL の倍率(factor)確定専用（他の op は使わない・
       既定値のままで後方互換）。倍率の出典は resolved["_sources"]["factor"] に、
       LLM 由来の値との食い違いは resolved["_warnings"] に積む（戻り値のタプル形は
       変えない＝呼び出し側の unpack を壊さない）。"""
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
        # ★ W3: target が実在しない場合、翻訳が「新しい列の名前」（例:「利益列を作って」の
        #   『利益』）を target と誤って埋めていることが多いと実測された（qwen2.5-coder:7b が
        #   『既存列に書く/新規に作る』の区別を安定して守らない）。実在しない＝一意に決まらない
        #   （複数解釈で曖昧）のとは別の理由なので、その場合だけ target を無指定として扱い
        #   新規列作成にフォールバックする（推測で断定しない CLARIFY の原則は、真に曖昧な
        #   ケース＝digit_candidates の複数一致にだけ残す）。
        if resolved.get("target"):
            v, was_inferred, err = resolve_col_ref(resolved["target"], headers.get(first_sheet, []))
            if err:
                if "一意に決まりません" in err:
                    return False, resolved, inferred, err
                del resolved["target"]
            else:
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

    elif op == "APPEND_TOTAL":
        if (err := resolve_in("col", first_sheet)):
            return False, resolved, inferred, err
        # ★ W6: label は既定値を持つ任意項目。ここで確定させ、codegen/事後条件/
        #   確認行の全部に同じ既定解決を一貫して渡す。
        resolved["label"] = str(resolved.get("label") or "合計")
        label = resolved["label"]

        # ★ A': factor は LLM から受け取らない。LLM が返した値(あれば)はいったん取り出して
        #   おき、機械抽出/用語集の結果と食い違う場合だけ WARN として記録する（常に機械が勝つ）。
        llm_factor_raw = resolved.pop("factor", None)

        text_factor, text_snippet = extract_rate_factor(task)
        vocab_factor, vocab_term = (None, None)
        if text_factor is None:
            vocab_factor, vocab_term = lookup_vocab_factor(task, vocab or {})

        sources: dict = {}
        if text_factor is not None:
            resolved["factor"] = text_factor
            sources["factor"] = f"依頼文: {text_snippet}"
        elif vocab_factor is not None:
            resolved["factor"] = vocab_factor
            sources["factor"] = f"用語集: {vocab_term}"
        else:
            resolved["factor"] = 1.0

        if resolved["factor"] <= 0:
            return False, resolved, inferred, f"倍率『{resolved['factor']}』は正の数でなければなりません"

        # ★ 恒真式の番人（最優先）: label が「税」/「込」を含むのに倍率が確定できず既定
        #   1.0 のままだと、税抜き金額に「税込み」ラベルが付いた恒真の誤りを事後条件が
        #   pass にしてしまう（args 基準の検証だから）。ここで機械的に CLARIFY へ倒す
        #   （語リストは 税/込 の2語で凍結・むやみに増やさない）。
        if resolved["factor"] == 1.0 and any(k in label for k in ("税", "込")):
            return False, resolved, inferred, (
                f"ラベル『{label}』は税/込を含みますが倍率が分かりません。"
                "依頼文に税率を書く（例:「消費税10%」）か、用語集に登録してください"
                "（例: ailine vocab add 消費税 1.1）"
            )

        if sources:
            resolved["_sources"] = sources
        if llm_factor_raw not in (None, ""):
            try:
                llm_factor = float(llm_factor_raw)
            except (TypeError, ValueError):
                llm_factor = None
            if llm_factor is not None and abs(llm_factor - resolved["factor"]) > 1e-9:
                mfactor = resolved["factor"]
                resolved["_warnings"] = [
                    f"LLM が返した倍率({llm_factor:g})と機械抽出の倍率({mfactor:g})が"
                    f"食い違うため機械抽出({mfactor:g})を採用しました"
                ]

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
    "APPEND_TOTAL": (("対象列", "col", None), ("ラベル", "label", None), ("倍率", "factor", None)),
}


def format_confirmation_line(op: str, resolved_args: dict, inferred: set) -> str:
    """命令言語形式の確認行を1行で組む（例: 解釈: 操作:並べ替え 対象:金額 順:降順）。
       推定で埋めた（数字表記から解決した等）引数には (推定) を付ける。
       ★ M2c: キー自体が resolved_args に無い任意項目（COMPUTE_COLUMN の target 等）は
       そのフィールドを丸ごと省略する（必須項目は常に存在するので既存の表示は変わらない）。
       ★ A': resolved_args["_sources"] に該当キーの出典があれば（例: 倍率:1.1）
       末尾に「（用語集: 消費税）」のように出典を添える（verify_dsl_args の APPEND_TOTAL が積む）。"""
    sources = resolved_args.get("_sources") or {}
    parts = [f"操作:{OP_LABELS.get(op, op)}"]
    for label, key, transform in _CONFIRM_FIELDS.get(op, ()):
        if key not in resolved_args:
            continue
        val = resolved_args.get(key)
        shown = transform(val) if transform else val
        tag = "(推定)" if key in inferred else ""
        if key in sources:
            tag += f"（{sources[key]}）"
        parts.append(f"{label}:{shown}{tag}")
    return "解釈: " + " ".join(parts)


# --- ④ 決定論 codegen（op → Basic テンプレ。LLM は一切使わない） --------------

def _wrap_basic(body: str) -> str:
    """CONTRACT と同じ骨格（Option 2行 + Sub Run(oDoc As Object) 1つ）で包む。"""
    return "Option VBASupport 1\nOption Explicit\n\nSub Run(oDoc As Object)\n" + body + "End Sub\n"


def _scan_last_row_basic(var: str = "oSheet", key_col: str = "0",
                          start_row: str = "1", min_ok: str | None = None) -> str:
    """走査ループの定型（refs の作法どおり：A列を上から走査して最終データ行を探す）。
       ★ W3: start_row（Basic 0起点の走査開始行＝見出し行の直下）を渡すと、見出しが
       物理1行目でない帳票でも正しい行から走査する（既定 "1" は旧挙動と同一）。
       min_ok（"lastRow がこの値未満なら Exit Sub"の閾値）を渡すと保存境界を調整できる
       （既定は start_row と同じ＝『データ0行なら何もしない』。見出しを含めて範囲を
       スタイリングする操作は min_ok に start_row-1 相当を渡す＝データ0行でも見出しは扱える）。"""
    if min_ok is None:
        min_ok = start_row
    return (f"    lastRow = {start_row}\n"
            f"    Do While {var}.getCellByPosition({key_col}, lastRow).getString() <> \"\"\n"
            f"        lastRow = lastRow + 1\n"
            f"    Loop\n"
            f"    lastRow = lastRow - 1\n"
            f"    If lastRow < {min_ok} Then Exit Sub\n")


def codegen_dsl(op: str, resolved_args: dict, book_meta: dict, use_formula: bool = True) -> str:
    """④ 決定論 codegen。既存ヘルパへの Call を最優先し、無い操作だけテンプレ Basic を書く。
       ★ W3: book_meta["header_rows"] があれば1枚目シートの見出し行(1起点)をそこから読み、
       0起点(hr0)に変換して全 op に一貫して渡す（『三層全部が同じ見出し推定を使う』の codegen 側）。
       header_rows が無い/キーが無い book_meta（_SAMPLE_META 等の旧テスト値）は既定1行目
       （hr0=0）＝旧挙動と完全に同一の Basic を生成する。
       use_formula: COMPUTE_COLUMN の既定を式（=B2*C2 等）にする（★ W3 Part3）。False で
       従来の値ベタ書きに戻す（--values）。"""
    headers = book_meta["headers"]
    first_sheet = book_meta["sheets"][0]
    header_row = book_meta.get("header_rows", {}).get(first_sheet, 1)
    hr0 = header_row - 1   # Basic 0起点の見出し行

    if op == "SORT":
        col_idx = headers[first_sheet].index(resolved_args["col"])
        asc = "True" if resolved_args["order"] == "asc" else "False"
        last_col = len(headers[first_sheet]) - 1
        return _wrap_basic(f"    Call SortByColumn(oDoc, {hr0}, {last_col}, {col_idx}, {asc})\n")

    if op == "LOOKUP_FILL":
        theaders = headers[resolved_args["target_sheet"]]
        key_idx = theaders.index(resolved_args["key_col"])
        tgt_idx = theaders.index(resolved_args["target_col"])
        src = resolved_args["source_sheet"].replace('"', '""')
        return _wrap_basic(f'    Call VLookupFromTable(oDoc, {hr0}, {key_idx}, {tgt_idx}, "{src}")\n')

    if op == "AGGREGATE":
        g_idx = headers[first_sheet].index(resolved_args["group_col"])
        v_idx = headers[first_sheet].index(resolved_args["value_col"])
        return _wrap_basic(f"    Call SummaryTable(oDoc, {hr0}, {g_idx}, {v_idx})\n")

    if op == "NUMBER_FORMAT":
        col_idx = headers[first_sheet].index(resolved_args["col"])
        return _wrap_basic(f"    Call FormatThousands(oDoc, {hr0}, {col_idx})\n")

    if op == "MERGE":
        c1s, r1s, c2s, r2s = re.match(
            r"([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)", resolved_args["range"]).groups()
        col1 = column_index_from_string(c1s.upper()) - 1
        col2 = column_index_from_string(c2s.upper()) - 1
        row1, row2 = int(r1s) - 1, int(r2s) - 1
        return _wrap_basic(f"    Call MergeCells(oDoc, {col1}, {row1}, {col2}, {row2})\n")

    if op == "CHART":
        v_idx = headers[first_sheet].index(resolved_args["value_col"])
        return _wrap_basic(f"    Call InsertBarChart(oDoc, {hr0}, {v_idx})\n")

    if op == "APPEND_TOTAL":
        # ★ W6: ヘルパは無し（罫線・カンマ等の見栄えまでは踏み込まない・素の SUM 式だけ）。
        #   データ最終行の直下に [ラベル文字列 | 合計式] を書く。
        #   ラベルは対象列の左隣に置く（既存の帳票では『合計』の文字が金額の左に来るのが自然
        #   で、対象列自体を上書きしない＝既存構造を壊さない置き方）。対象列が表の最左端
        #   （col_idx=0）の場合は左隣が無いため、値のみを書きラベルは省略する。
        # ★ B: 挿入耐性式（bench/formula_spike_work2 で実測）。SUM(D2:INDEX(D:D;ROW()-1))
        #   型で書く。INDEX(D:D;ROW()-1) は「この式自身の1行上」を指すので、後でデータ行を
        #   1本挿入しても SUM 範囲が自動で追従する（静的な "D2:D5" 型は追従しない）。
        #   ★ setFormula は LO 方言＝INDEX の2引数区切りはセミコロン(;)。カンマ(,)で
        #   書くと #VALUE!/#NAME? になる（実測: formula_spike_RESULTS.md 追記分）。
        #   保存後は自動でカンマ形 "=SUM(D2:INDEX(D:D,ROW()-1))*1.1" に変換される
        #   （check_append_total はこの保存後カンマ形と照合する）。
        col_idx = headers[first_sheet].index(resolved_args["col"])
        label = str(resolved_args.get("label", "合計")).replace('"', '""')
        factor = float(resolved_args.get("factor", 1) or 1)
        col_letter = get_column_letter(col_idx + 1)
        start_excel_row = hr0 + 2   # データ先頭行（Basic 0起点 hr0+1）の Excel(1起点) 行
        factor_tail = "" if factor == 1 else f"*{factor:g}"
        body = ("    Dim oSheet As Object, lastRow As Long, totalRow As Long\n"
                "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                + _scan_last_row_basic(start_row=str(hr0 + 1))
                + "    totalRow = lastRow + 1\n")
        if col_idx > 0:
            body += f'    oSheet.getCellByPosition({col_idx - 1}, totalRow).setString("{label}")\n'
        body += (f'    oSheet.getCellByPosition({col_idx}, totalRow).setFormula('
                 f'"=SUM(" & "{col_letter}" & {start_excel_row} & ":INDEX(" & "{col_letter}" & '
                 f'":" & "{col_letter}" & ";ROW()-1))" & "{factor_tail}")\n')
        return _wrap_basic(body)

    if op == "CENTER_ALIGN":
        if resolved_args["target"] == "all":
            last_col = len(headers[first_sheet]) - 1
            return _wrap_basic(f"    Call AlignCenter(oDoc, {hr0}, {last_col})\n")
        # col:NAME はヘルパ無し → refs の作法（走査して範囲を求め HoriJustify）でテンプレを書く。
        col_idx = headers[first_sheet].index(resolved_args["target"][4:])
        body = ("    Dim oSheet As Object, oRange As Object, lastRow As Long\n"
                "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                + _scan_last_row_basic(start_row=str(hr0 + 1), min_ok=str(hr0))
                + f"    oRange = oSheet.getCellRangeByPosition({col_idx}, {hr0}, {col_idx}, lastRow)\n"
                "    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER\n")
        return _wrap_basic(body)

    if op == "BOLD":
        target = resolved_args["target"]
        if target.startswith("row:"):
            row_idx = int(target[4:]) - 1
            # ★ W3: 列幅は Basic で走査せず、接地済みの見出し列数(headers)から決定論的に
            #   決める（多段見出しで先頭列が空欄のケースで走査が誤って -1 になるのを回避）。
            last_col = len(headers[first_sheet]) - 1
            body = (f"    Call StyleBold(oDoc, 0, {row_idx}, {last_col}, {row_idx})\n")
        else:
            col_idx = headers[first_sheet].index(target[4:])
            body = ("    Dim oSheet As Object, lastRow As Long\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    + _scan_last_row_basic(start_row=str(hr0 + 1), min_ok=str(hr0))
                    + f"    Call StyleBold(oDoc, {col_idx}, {hr0}, {col_idx}, lastRow)\n")
        return _wrap_basic(body)

    if op == "FILL_COLOR":
        target = resolved_args["target"]
        hexcolor = COLOR_MAP[resolved_args["color"]]
        if target.startswith("row:"):
            row_idx = int(target[4:]) - 1
            # ★ W3: 列幅は Basic で走査せず、接地済みの見出し列数(headers)から決定論的に
            #   決める（多段見出しで先頭列が空欄のケースで走査が誤って -1 になるのを回避）。
            last_col = len(headers[first_sheet]) - 1
            body = ("    Dim oSheet As Object, c As Integer\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    f"    For c = 0 To {last_col}\n"
                    f"        oSheet.getCellByPosition(c, {row_idx}).CellBackColor = &H{hexcolor}&\n"
                    "    Next c\n")
        else:
            col_idx = headers[first_sheet].index(target[4:])
            body = ("    Dim oSheet As Object, lastRow As Long, r As Long\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    + _scan_last_row_basic(start_row=str(hr0 + 1), min_ok=str(hr0))
                    + f"    For r = {hr0} To lastRow\n"
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
            header_write = f"    oSheet.getCellByPosition({new_col}, {hr0}).setString(\"{header_name}\")\n"
        # ★ W3 Part3: 既定は式（setFormula）。formula_spike の実測どおり、単純な行内の
        #   二項演算（=B2*C2 型）は区切り記号もシート参照も無いため LO 方言(;/.)の
        #   気遣いが不要（多引数関数・シート跨ぎ参照だけが方言の対象＝ここでは無縁）。
        #   行番号だけが Basic 側ループ変数(i)で動くので、その部分だけ実行時に連結する。
        if use_formula:
            col1_letter = get_column_letter(i1 + 1)
            col2_letter = get_column_letter(i2 + 1)
            write_line = (f'        oSheet.getCellByPosition({new_col}, i).setFormula('
                          f'"=" & "{col1_letter}" & (i + 1) & "{operator}" & "{col2_letter}" & (i + 1))\n')
        else:
            write_line = (f"        oSheet.getCellByPosition({new_col}, i).setValue("
                          f"oSheet.getCellByPosition({i1}, i).getValue() {operator} "
                          f"oSheet.getCellByPosition({i2}, i).getValue())\n")
        body = ("    Dim oSheet As Object, lastRow As Long, i As Long\n"
                "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                + _scan_last_row_basic(start_row=str(hr0 + 1))
                + header_write
                + f"    For i = {hr0 + 1} To lastRow\n"
                + write_line
                + "    Next i\n")
        return _wrap_basic(body)

    raise ValueError(f"未対応の op: {op}")


# --- ⑥ op 別事後条件（達成の機械検証。openpyxl で out ファイルを読むだけ・LO 不要） ----

def _col_index_by_header(ws, name: str, header_row: int = 1):
    """見出し行(既定は物理1行目・header_row で指定可)を左から走査して name に一致する
       列の1起点インデックスを返す。無ければ None。★ W3: header_row 省略時は旧挙動と同一。
       ★ W3: header_row>1 の子見出し行で空欄の列は、book_columns() と同じ規則で
       真上の行を遡って引き継ぐ（多段見出しの先頭列対策。無いと『7月』等キー列より
       手前の空欄列で走査が誤って打ち切られる）。"""
    c = 1
    while True:
        v = ws.cell(row=header_row, column=c).value
        if v in (None, "") and header_row > 1:
            for up in range(header_row - 1, 0, -1):
                uv = ws.cell(row=up, column=c).value
                if uv not in (None, ""):
                    v = uv
                    break
        if v in (None, ""):
            return None
        if str(v) == name:
            return c
        c += 1


def _scan_last_row(ws, key_col: int = 1, header_row: int = 1) -> int:
    """key_col(1起点)を上から走査した最終データ行（見出し行を除く）。データが無ければ
       header_row。★ W3: header_row 省略時は旧挙動（見出し=1行目・データ開始=2行目）と同一。"""
    r = header_row + 1
    while ws.cell(row=r, column=key_col).value not in (None, ""):
        r += 1
    return r - 1


def _scan_last_col(ws, header_row: int = 1) -> int:
    """見出し行(既定は物理1行目)を左から走査した最終列（1起点）。
       ★ W3: _col_index_by_header と同じ規則で、子見出し行(header_row>1)の空欄列は
       真上の行を遡って引き継ぐ（多段見出しの先頭列対策）。"""
    def _effective(c: int):
        v = ws.cell(row=header_row, column=c).value
        if v in (None, "") and header_row > 1:
            for up in range(header_row - 1, 0, -1):
                uv = ws.cell(row=up, column=c).value
                if uv not in (None, ""):
                    return uv
        return v
    c = 1
    while _effective(c) not in (None, ""):
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


def _is_number(v) -> bool:
    """★ 止血2: bool は int のサブクラスだが数値セルとしては扱わない（True/False混入対策）。"""
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ★ 止血1/2 共通の文言。事後条件チェッカーは検証対象0件を絶対に「合格」にしない
#   （D検体: no-opを『行数が少なく比較不要』で素通ししていた根治）。
_ZERO_TARGET_REASON = "事後条件の検証対象が0件（何も検証できていない）"


def check_sort(path: Path, args: dict, header_row: int = 1) -> tuple:
    """SORT の事後条件。戻り値は (status, reason)。status ∈ {"pass","warn","fail"}。
       ★ 止血1: 検証対象が0件なら fail、1件（順序が定義できない）なら warn とし、
       どちらも「機械検証済み」とは名乗らない。
       ★ 止血2: 合計行等の非数値/None セルは比較から除外し、除外件数を表示する
       （C②: None >= int の生トレースバックの根治）。全部除外なら0件と同じ扱い。
       ★ W3: header_row(1起点、省略時1) が『接地・codegen』と同じ見出し行を指す。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    idx = _col_index_by_header(ws, args["col"], header_row=header_row)
    if idx is None:
        wb.close()
        return "fail", f"列『{args['col']}』が見つからない"
    last = _scan_last_row(ws, header_row=header_row)
    raw_vals = [ws.cell(row=r, column=idx).value for r in range(header_row + 1, last + 1)]
    wb.close()
    vals = [v for v in raw_vals if _is_number(v)]
    excluded = len(raw_vals) - len(vals)
    note = f"（数値でない {excluded} 行は対象外）" if excluded else ""
    if len(vals) == 0:
        return "fail", _ZERO_TARGET_REASON + note
    if len(vals) == 1:
        return "warn", f"検証対象が1行のみ（並べ替えの意味がありません）{note}"
    asc = args["order"] == "asc"
    ok = (all(vals[i] <= vals[i + 1] for i in range(len(vals) - 1)) if asc
          else all(vals[i] >= vals[i + 1] for i in range(len(vals) - 1)))
    if not ok:
        return "fail", f"列『{args['col']}』が指定順（{args['order']}）に並んでいない{note}"
    return "pass", f"{len(vals)} 行を検証（{'昇順' if asc else '降順'}）{note}"


def check_compute_column(path: Path, args: dict, header_row: int = 1,
                          use_formula: bool = False) -> tuple:
    """★ 止血1/2: 演算対象が非数値/None の行（合計行等）は「対象外」として除外し件数を
       表示する。除外後に検証できた行が0件なら fail（機械検証済みと名乗らない）。
       対象列(target)自体が非数値なのは演算が本当に効いていない証拠なので除外せず fail。
       ★ W3 Part3: use_formula=True のとき事後条件を二層化する — ①通常の openpyxl 読み
       で保存された式文字列が期待形か ②data_only 読みでキャッシュ値が演算結果と一致するか。
       両方合格して初めて pass にする（式だけ合っていて未計算/値だけ合っていて式が
       無いケースの両方を見逃さない）。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    op1, op2 = args["operands"]
    i1 = _col_index_by_header(ws, op1, header_row=header_row)
    i2 = _col_index_by_header(ws, op2, header_row=header_row)
    # ★ M2c: target(実在列名) 指定時はその列を検証する。無指定なら従来どおり自動命名の新列。
    target = args.get("target")
    newname = target or f"{op1}{args['operator']}{op2}"
    inew = _col_index_by_header(ws, newname, header_row=header_row)
    if i1 is None or i2 is None or inew is None:
        wb.close()
        return "fail", f"演算対象または対象列『{newname}』が見つからない"
    last = _scan_last_row(ws, header_row=header_row)
    wb_v = None
    ws_v = None
    if use_formula:
        wb_v = openpyxl.load_workbook(path, data_only=True)
        ws_v = wb_v[wb_v.sheetnames[0]]
    col1_letter = get_column_letter(i1)
    col2_letter = get_column_letter(i2)
    checked = 0
    excluded = 0
    for r in range(header_row + 1, last + 1):
        a = ws.cell(row=r, column=i1).value
        b = ws.cell(row=r, column=i2).value
        got = ws.cell(row=r, column=inew).value
        if not _is_number(a) or not _is_number(b):
            excluded += 1   # 例: 合計行で演算対象セルが空欄
            continue
        want = _apply_operator(a, b, args["operator"])
        if use_formula:
            expect_formula = f"={col1_letter}{r}{args['operator']}{col2_letter}{r}"
            if not isinstance(got, str) or got.replace(" ", "") != expect_formula:
                wb.close(); wb_v.close()
                return "fail", f"{r}行目: 式が期待形でない (期待 {expect_formula} 実際 {got!r})"
            got_cached = ws_v.cell(row=r, column=inew).value
            if not _is_number(got_cached) or abs(got_cached - want) > 1e-6:
                wb.close(); wb_v.close()
                return "fail", f"{r}行目: 式のキャッシュ値が不一致 (期待 {want} 実際 {got_cached!r})"
        else:
            if not _is_number(got) or abs(got - want) > 1e-6:
                wb.close()
                return "fail", f"{r}行目: 期待 {want} 実際 {got}"
        checked += 1
    wb.close()
    if wb_v is not None:
        wb_v.close()
    note = f"（数値でない {excluded} 行は対象外）" if excluded else ""
    if checked == 0:
        return "fail", _ZERO_TARGET_REASON + note
    if use_formula:
        return "pass", f"{checked} 行を検証（式・キャッシュ値とも一致）{note}"
    return "pass", f"{checked} 行を検証{note}"


def check_lookup_fill(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W3: header_row は対象シート(target_sheet=1枚目)の見出し行。参照表(source_sheet)は
       常に「列0=キー・列1=値」の物理1行目見出し前提（VLookupFromTable ヘルパの仕様どおり・
       検出対象外）。"""
    wb = openpyxl.load_workbook(path)
    if args["target_sheet"] not in wb.sheetnames or args["source_sheet"] not in wb.sheetnames:
        wb.close()
        return "fail", "対象/参照シートが無い"
    tws = wb[args["target_sheet"]]
    sws = wb[args["source_sheet"]]
    key_idx = _col_index_by_header(tws, args["key_col"], header_row=header_row)
    tgt_idx = _col_index_by_header(tws, args["target_col"], header_row=header_row)
    if key_idx is None or tgt_idx is None:
        wb.close()
        return "fail", "対象シートにキー列/対象列が無い"
    lookup = {}
    r = 2
    while sws.cell(row=r, column=1).value not in (None, ""):
        lookup[sws.cell(row=r, column=1).value] = sws.cell(row=r, column=2).value
        r += 1
    scanned = 0   # ★ 止血1: 対象シートに行が1件も無い(0件)場合と、行はあるが1件も
                  #   対応表に載っていない場合を別のメッセージで区別する。
    checked = 0
    r = header_row + 1
    while tws.cell(row=r, column=key_idx).value not in (None, ""):
        scanned += 1
        key = tws.cell(row=r, column=key_idx).value
        if key in lookup:
            got = tws.cell(row=r, column=tgt_idx).value
            want = lookup[key]
            if got != want:
                wb.close()
                return "fail", f"{r}行目: キー『{key}』の転記値が不一致 (期待 {want!r} 実際 {got!r})"
            checked += 1
        r += 1
    wb.close()
    if scanned == 0:
        return "fail", _ZERO_TARGET_REASON
    if checked == 0:
        return "fail", "対応表に載っているキーが1件も転記されていない"
    return "pass", f"{checked} 行を検証"


def check_aggregate(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W3: header_row は集計元(1枚目)の見出し行。出力の「集計」シートは SummaryTable
       ヘルパが毎回新規作成し常に物理1行目が見出し（検出対象外・そのまま）。"""
    wb = openpyxl.load_workbook(path)
    if "集計" not in wb.sheetnames:
        wb.close()
        return "fail", "『集計』シートが無い"
    src = wb[wb.sheetnames[0]]
    gi = _col_index_by_header(src, args["group_col"], header_row=header_row)
    vi = _col_index_by_header(src, args["value_col"], header_row=header_row)
    if gi is None or vi is None:
        wb.close()
        return "fail", "分類列/集計列が見つからない"
    expect: dict = {}
    r = header_row + 1
    while src.cell(row=r, column=1).value not in (None, ""):
        k = src.cell(row=r, column=gi).value
        v = src.cell(row=r, column=vi).value
        v = v if _is_number(v) else 0   # ★ 止血2: 非数値/None は0扱い（クラッシュさせない）
        expect[k] = expect.get(k, 0) + v
        r += 1
    if not expect:
        wb.close()
        return "fail", _ZERO_TARGET_REASON   # ★ 止血1: 集計元データが0件を「合格」にしない
    out = wb["集計"]
    seen = set()
    r = 2
    while True:
        k = out.cell(row=r, column=1).value
        if k in (None, "") or k == "合計":
            break
        v = out.cell(row=r, column=2).value
        v = v if _is_number(v) else 0
        if k not in expect or abs(v - expect[k]) > 1e-6:
            wb.close()
            return "fail", f"グループ『{k}』の合計が不一致 (期待 {expect.get(k)} 実際 {v})"
        seen.add(k)
        r += 1
    wb.close()
    if seen != set(expect.keys()):
        return "fail", "集計に含まれないグループがある"
    return "pass", f"{len(expect)} グループを検証"


def check_bold(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W3: "col:" 対象は見出し(header_row)を含めて検証する（codegen の
       StyleBold(oDoc, col, hr0, col, lastRow) が見出しも含めて太字にするため）。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    kind, val = args["target"].split(":", 1)
    if kind == "row":
        last_col = _scan_last_col(ws, header_row=header_row)
        row = int(val)
        cells = [ws.cell(row=row, column=c) for c in range(1, last_col + 1)]
        label = f"{row}行目"
    else:
        idx = _col_index_by_header(ws, val, header_row=header_row)
        if idx is None:
            wb.close()
            return "fail", f"列『{val}』が見つからない"
        last_row = _scan_last_row(ws, header_row=header_row)
        cells = [ws.cell(row=r, column=idx) for r in range(header_row, last_row + 1)]
        label = f"列『{val}』"
    if not cells:   # ★ 止血1: 検証対象0件（見出しすら無い空シート等）を合格にしない
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    ok = all(c.font and c.font.bold for c in cells)
    wb.close()
    if not ok:
        return "fail", f"{label} に太字でないセルがある"
    return "pass", f"{len(cells)} セルが太字"


def check_fill_color(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W3: "col:" 対象は見出し(header_row)を含めて検証する（codegen が見出しも
       含めて塗るため）。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    want_hex = COLOR_MAP[args["color"]].upper()
    kind, val = args["target"].split(":", 1)
    if kind == "row":
        last_col = _scan_last_col(ws, header_row=header_row)
        row = int(val)
        cells = [ws.cell(row=row, column=c) for c in range(1, last_col + 1)]
        label = f"{row}行目"
    else:
        idx = _col_index_by_header(ws, val, header_row=header_row)
        if idx is None:
            wb.close()
            return "fail", f"列『{val}』が見つからない"
        last_row = _scan_last_row(ws, header_row=header_row)
        cells = [ws.cell(row=r, column=idx) for r in range(header_row, last_row + 1)]
        label = f"列『{val}』"
    if not cells:   # ★ 止血1
        wb.close()
        return "fail", _ZERO_TARGET_REASON

    def _matches(cell) -> bool:
        if cell.fill is None or not cell.fill.patternType:
            return False
        return str(cell.fill.start_color.rgb).upper().endswith(want_hex)

    ok = all(_matches(c) for c in cells)
    wb.close()
    if not ok:
        return "fail", f"{label} に色『{args['color']}』が付いていないセルがある"
    return "pass", f"{len(cells)} セルの背景色を確認"


def check_number_format(path: Path, args: dict, header_row: int = 1) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    idx = _col_index_by_header(ws, args["col"], header_row=header_row)
    if idx is None:
        wb.close()
        return "fail", f"列『{args['col']}』が見つからない"
    last = _scan_last_row(ws, header_row=header_row)
    if last < header_row + 1:   # ★ 止血1: データ行0件を合格にしない
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    ok = all("#,##0" in (ws.cell(row=r, column=idx).number_format or "")
             for r in range(header_row + 1, last + 1))
    wb.close()
    if not ok:
        return "fail", f"列『{args['col']}』に桁区切り書式が付いていないセルがある"
    return "pass", f"{last - header_row} 行に桁区切り書式を確認"


def check_merge(path: Path, args: dict, header_row: int = 1) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    ranges = {str(r) for r in ws.merged_cells.ranges}
    wb.close()
    if args["range"] not in ranges:
        return "fail", f"範囲『{args['range']}』が結合されていない"
    return "pass", f"{args['range']} の結合を確認"


def check_chart(path: Path, before_charts: int) -> tuple:
    after = _charts_count(path)
    if after != before_charts + 1:
        return "fail", f"グラフ数が +1 でない（{before_charts} → {after}）"
    return "pass", f"グラフ数 {before_charts} → {after}"


def check_center_align(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W3: "all"/"col:" とも見出し(header_row)を含めて検証する（codegen の
       AlignCenter/inline テンプレが見出しも含めて中央揃えにするため）。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    target = args["target"]
    if target == "all":
        last_row = _scan_last_row(ws, header_row=header_row)
        last_col = _scan_last_col(ws, header_row=header_row)
        cells = [ws.cell(row=r, column=c) for r in range(header_row, last_row + 1)
                 for c in range(1, last_col + 1)]
        label = "表全体"
    else:
        colname = target[4:]
        idx = _col_index_by_header(ws, colname, header_row=header_row)
        if idx is None:
            wb.close()
            return "fail", f"列『{colname}』が見つからない"
        last_row = _scan_last_row(ws, header_row=header_row)
        cells = [ws.cell(row=r, column=idx) for r in range(header_row, last_row + 1)]
        label = f"列『{colname}』"
    if not cells:   # ★ 止血1
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    ok = all(c.alignment and c.alignment.horizontal == "center" for c in cells)
    wb.close()
    if not ok:
        return "fail", f"{label} に中央揃えでないセルがある"
    return "pass", f"{len(cells)} セルの中央揃えを確認"


_APPEND_TOTAL_FORMULA_RE = re.compile(
    r"^=SUM\([A-Za-z]{1,3}\d+:INDEX\([A-Za-z]{1,3}:[A-Za-z]{1,3},ROW\(\)-1\)\)(\*[\d.]+)?$",
    re.IGNORECASE)


def check_append_total(path: Path, args: dict, header_row: int = 1) -> tuple:
    """★ W6: APPEND_TOTAL の事後条件（二層・COMPUTE_COLUMN の use_formula 版と同じ考え方）。
       ①式文字列が期待形『=SUM(<列><行>:INDEX(<列>:<列>,ROW()-1))』（★ B: 挿入耐性式・
       factor!=1 のときは末尾に *factor が付く）②data_only 読みのキャッシュ値が
       「列合計×factor」と一致（浮動小数は丸め許容）。対象列が表の最左端でない場合は、
       その左隣セルに期待ラベルが立っていることも確認する（codegen_dsl の配置と対）。
       ★ B: codegen は LO 方言（INDEX の引数区切りがセミコロン）で書くが、保存後は
       カンマ区切りに自動変換される（実測）ので、ここではその保存後のカンマ形と照合する。

       ★ 合計行の位置は『対象列そのものに "=SUM(" で始まる式が最初に現れた行』として直接
       探す（生読み・型判定はしない）。実測で分かった罠が2つあり、どちらもこの探し方だけ
       で同時に避けられる:
       - COMPUTE_COLUMN の式化（W3 Part3・既定）で対象列自体が "=B2*C2" 型の式で埋まって
         いることがある。生セル値の型（数値かどうか）だけで『データ行/合計行』を分けようと
         すると、データ行の式もキャッシュ値未読の合計行の式も同じ『文字列』に見えて区別
         できない（E2E 実測: 小計列を先に式化してから APPEND_TOTAL するケースで発覚）。
       - 最終データ行の判定に列1(左端)を走査する既存の _scan_last_row は使わない。左端が
         分類列で、かつ対象列がその右隣（＝ラベルの置き場所が偶然その分類列と重なる、例:
         2列だけの帳票）だと、書き込んだラベルの文字列自体が『データ行』として誤って
         数えられ off-by-one になる（実測）。
       "=SUM(" という固有の目印を対象列自身の中だけで探すので、他列の中身にも
       COMPUTE_COLUMN の式の形にも影響されない。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    idx = _col_index_by_header(ws, args["col"], header_row=header_row)
    if idx is None:
        wb.close()
        return "fail", f"列『{args['col']}』が見つからない"
    r = header_row + 1
    while True:
        v = ws.cell(row=r, column=idx).value
        if v is None:
            break
        if isinstance(v, str) and v.replace(" ", "").startswith("=SUM("):
            break
        r += 1
    last = r - 1
    if last < header_row + 1:
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    total_row = r
    got_formula = ws.cell(row=total_row, column=idx).value
    got_formula_norm = got_formula.replace(" ", "") if isinstance(got_formula, str) else ""
    if not _APPEND_TOTAL_FORMULA_RE.match(got_formula_norm):
        wb.close()
        detail = f"{total_row}行目: 合計の式が期待形(挿入耐性 SUM 型)でない (実際 {got_formula!r})"
        return "fail", detail
    label_ok = True
    want_label = str(args.get("label") or "合計")
    got_label = None
    if idx > 1:
        got_label = ws.cell(row=total_row, column=idx - 1).value
        label_ok = got_label == want_label
    wb.close()
    if not label_ok:
        return "fail", f"{total_row}行目: ラベルが期待『{want_label}』と不一致 (実際 {got_label!r})"

    wb_v = openpyxl.load_workbook(path, data_only=True)
    ws_v = wb_v[wb_v.sheetnames[0]]
    raw_vals = [ws_v.cell(row=rr, column=idx).value for rr in range(header_row + 1, last + 1)]
    nums = [v for v in raw_vals if _is_number(v)]
    got_cached = ws_v.cell(row=total_row, column=idx).value
    wb_v.close()
    if not nums:
        return "fail", _ZERO_TARGET_REASON
    factor = float(args.get("factor", 1) or 1)
    want_total = sum(nums) * factor
    if not _is_number(got_cached) or abs(got_cached - want_total) > 1e-6:
        return "fail", f"{total_row}行目: 合計のキャッシュ値が不一致 (期待 {want_total} 実際 {got_cached!r})"
    return "pass", f"{len(nums)} 行の合計を検証（式・キャッシュ値・ラベルとも一致）"


POSTCONDITIONS = {
    "SORT": check_sort, "COMPUTE_COLUMN": check_compute_column,
    "LOOKUP_FILL": check_lookup_fill, "AGGREGATE": check_aggregate,
    "BOLD": check_bold, "FILL_COLOR": check_fill_color,
    "NUMBER_FORMAT": check_number_format, "MERGE": check_merge,
    "CENTER_ALIGN": check_center_align, "APPEND_TOTAL": check_append_total,
}


def run_postcondition(op: str, out_book: Path, resolved_args: dict, before_charts: int = 0,
                       header_row: int = 1, use_formula: bool = False) -> tuple:
    """⑥ op 別事後条件。(status, reason)。status ∈ {"pass","warn","fail","error"}。
       CHART だけ before_charts と比較する専用の形。
       ★ W3: header_row（1起点、省略時1）を全チェッカーに一貫して渡す（『三層全部が
       同じ見出し推定を使う』の事後条件側）。use_formula は COMPUTE_COLUMN 専用（W3 Part3）。
       ★ 止血2: チェッカー内で予期しない例外が起きても生の Python トレースバックを
       出さない。ここで必ず捕まえて "error" ステータス + 要約1行に変換する
       （C②の教訓: 事後条件チェッカー自身のクラッシュがユーザーに未捕捉のまま漏れていた）。"""
    try:
        if op == "CHART":
            return check_chart(out_book, before_charts)
        fn = POSTCONDITIONS.get(op)
        if fn is None:
            return "fail", f"未対応の op: {op}"
        if op == "COMPUTE_COLUMN":
            return fn(out_book, resolved_args, header_row, use_formula)
        return fn(out_book, resolved_args, header_row)
    except Exception as e:
        return "error", f"事後条件の検証に失敗: {type(e).__name__}: {e}"


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
#  ★ W3: 正規化パスは NOOP_MACRO でなく _structdump_macro() を使う（同じ LO 往復で
#     構造読み取りを二役させる）。文書へは何も書かない点は NOOP_MACRO と同じ。


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
    dump_path = workdir / STRUCTDUMP_FILENAME
    struct_code = _structdump_macro(dump_path)
    ok, err, _ = basrun_apply(normalized, struct_code, workdir, timeout=timeout)
    if not ok:
        _stop_office()
        shutil.copy2(book, normalized)   # 中途半端な保存状態を残さず作り直す
        ok, err, _ = basrun_apply(normalized, struct_code, workdir, timeout=timeout)
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
        # ★ A': 用語集/依頼文から機械確定した値（APPEND_TOTAL の倍率等）の出典。無ければ None。
        "provenance": result.get("provenance"),
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


# --- ★ A': 用語集(vocab) コマンド --------------------------------------------

def cmd_vocab(a: argparse.Namespace) -> int:
    """`ailine vocab add <語> <値>` / `ailine vocab list`。
       ★ remove は作らない（今の実需では add/list の2本で足りる。壊れたら vocab.json を
       直接編集すればよい）。"""
    if a.vocab_cmd == "add":
        ok, msg = vocab_add(a.term, a.value)
        print(("✓ " if ok else "× ") + msg)
        return 0 if ok else 1
    # list
    vocab = load_vocab()
    if not vocab:
        print(f"（用語集は空。{VOCAB_FILE} に登録するか `ailine vocab add <語> <値>` で追加）")
        return 0
    print(f"用語集（{VOCAB_FILE}・{len(vocab)}件）:")
    for term in sorted(vocab):
        print(f"  {term} = {vocab[term]:g}")
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
    """run コマンドの入口。★ W3: 正規化パス(＋StructDump による見出し行推定)を
       翻訳より前に一度だけ行う（『三層全部が同じ見出し推定を使う』ための土台。
       translate_task 自身の接地(book_meta)も検出した見出し行を使う）。
       --dry は従来どおり LibreOffice に触れない（見出しは物理1行目のまま・E2E 対象外）。
       ① 見出し行推定（StructDump） → 自信不足なら CLARIFY して exit 3
       ② 翻訳（計画）→
       - 計画が空/1段で CLARIFY → 質問して exit 3
       - 計画が空/1段で DSL 語彙 → ③〜⑥の決定論パイプライン(cmd_run_dsl)
       - 計画が空/1段でそれ以外(FREEFORM・翻訳失敗) → 現行の自由生成経路(cmd_run_freeform)
       - 計画が2段以上(複合依頼) → 段ごとに honest な項目別実行(cmd_run_plan)（M2c）
       ★ 後方互換: translate_task が "plan" で包まない旧形式（bare {"op":...}）を返した場合
       （テストの monkeypatch を含む）も、その dict をそのまま単一段として扱う。"""
    book = Path(a.book).resolve()
    if not book.exists():
        sys.exit(f"文書が無い: {book}")

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    apply_timeout = a.timeout if a.timeout else None   # 0 で無効化（旧挙動 = 無制限）

    source_book = book
    struct_dump: dict = {}
    if not a.dry:
        t0 = progress_start("⏳ 初回準備（文書の正規化+構造読み取り）…")
        source_book = normalize_book(book, workdir, timeout=apply_timeout)
        progress_end(t0)
        struct_dump = build_struct_dump(source_book, workdir)

    sheets = build_book_meta(source_book).get("sheets", [])
    header_rows, clarify_q = resolve_header_rows(struct_dump, sheets)
    if clarify_q:
        print(f"？ {clarify_q}")
        return 3

    book_meta = build_book_meta(source_book, header_rows=header_rows)
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
            return cmd_run_dsl(a, book, source_book, book_meta, op, step.get("args", {}))
        return cmd_run_freeform(a, book, source_book)

    return cmd_run_plan(a, book, source_book, book_meta, plan)


def _column_has_existing_values(book_path: Path, sheet_name: str, col_name: str,
                                 header_row: int = 1) -> bool:
    """★ M2c: target(既存列指定)列に、見出し行を除いてどれか値が入っているか。
       上書き検知の明示用。読めない/列やシートが見つからない場合は False
       （保守的に『無い』扱い＝誤って警告しない）。★ W3: header_row(1起点)で見出しの
       実位置を受け取る（省略時は物理1行目・旧挙動と同一）。"""
    try:
        wb = openpyxl.load_workbook(book_path, read_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return False
        ws = wb[sheet_name]
        idx = _col_index_by_header(ws, col_name, header_row=header_row)
        if idx is None:
            wb.close()
            return False
        last = _scan_last_row(ws, header_row=header_row)
        found = any(ws.cell(row=r, column=idx).value not in (None, "")
                    for r in range(header_row + 1, last + 1))
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
    header_row = book_meta.get("header_rows", {}).get(sheets[0], 1)
    if _column_has_existing_values(book_path, sheets[0], resolved["target"], header_row=header_row):
        return f"★ 対象列『{resolved['target']}』には既存値があります（上書きします）"
    return None


def cmd_run_dsl(a: argparse.Namespace, book: Path, source_book: Path, book_meta: dict,
                 op: str, raw_args: dict) -> int:
    """M2b の決定論パイプライン本体。②検証 → ③確認行 → ④codegen → ⑤適用 → ⑥事後条件。
       ★ W3: source_book は cmd_run が翻訳より前に正規化済み（同じ LO 往復で StructDump も
       済ませてある）。ここでは正規化をやり直さない（二重 LO 起動を避ける）。
       book_meta["header_rows"] を codegen/事後条件へ一貫して渡す（三層とも同じ見出し推定）。"""
    vocab = load_vocab()
    ok, resolved, inferred, err = verify_dsl_args(op, raw_args, book_meta, task=a.task, vocab=vocab)
    if not ok:
        print(f"？ {err}")
        return 3

    first_sheet = book_meta["sheets"][0] if book_meta.get("sheets") else None
    header_row = book_meta.get("header_rows", {}).get(first_sheet, 1)
    use_formula = not getattr(a, "values", False)

    line = format_confirmation_line(op, resolved, inferred)
    print(f"■ ailine（DSL 経路）  model={a.model}  book={book.name}")
    print(line)
    warn_overwrite = _maybe_warn_target_overwrite(op, resolved, book_meta, book)
    if warn_overwrite:
        print(warn_overwrite)
    for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
        print(f"⚠ {w}")

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

    code = codegen_dsl(op, resolved, book_meta, use_formula=use_formula)
    (workdir / "dsl_attempt.bas").write_text(code, encoding="utf-8")
    print(f"\n─ 生成した .bas（決定論・LLM不使用）───────────────")
    print(code)
    print("──────────────────────────────────────────")

    result = {"ok": False, "attempts": 1, "task": a.task, "model": a.model,
              "path": "dsl", "command": line, "postcondition": None,
              "provenance": resolved.get("_sources")}

    if a.dry:
        print("（--dry: 適用しない。レビュー後に --dry を外して実行）")
        result["ok"] = True
        result["dry"] = True
        _finish_run(a, book, result, "none")
        return 0

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
    # ★ 止血3: DSL 経路は事後条件チェッカーが openpyxl で out ファイルを直接・全行
    #   読むため「表示だけ」が切り詰められている（適用・検証は全行）。
    notice = _truncation_notice(before, after, exhaustive_postcondition=True)
    if notice:
        print(notice)
    advisories = build_advisories(a.task, before, after)
    for adv in advisories:
        print(adv)
    result["changes"] = lines
    result["advisories"] = advisories

    status, reason = run_postcondition(op, out_book, resolved, before_charts=before["charts"],
                                        header_row=header_row, use_formula=use_formula)
    # ★ 止血1/2: status は "pass"/"warn"/"fail"/"error"。"error" はチェッカー内の
    #   予期しない例外を捕まえた印（--json 上は "fail" に丸める）。
    result["postcondition"] = "fail" if status == "error" else status
    if status == "error":
        print(f"\n× {reason}")
        result["out"] = str(out_book)
        _finish_run(a, book, result, "postcondition_error")
        return 1
    if status == "fail":
        print(f"\n× 適用されたが事後条件を満たさない: {reason}")
        result["out"] = str(out_book)
        _finish_run(a, book, result, "postcondition_fail")
        return 1
    if status == "warn":
        # ★ 止血1: 検証対象が少なすぎて意味を持たない場合、「機械検証済み」とは名乗らない。
        print(f"\n⚠ 事後条件を機械検証できなかった（操作:{OP_LABELS.get(op, op)}）: {reason}")
        result["ok"] = True
    else:
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


def cmd_run_freeform(a: argparse.Namespace, book: Path, source_book: Path) -> int:
    """自由生成経路（従来の cmd_run 本体そのまま。M2a の助言つき）。
       ① 翻訳が CLARIFY にも DSL 語彙にも決まらなかった（FREEFORM・翻訳失敗）ときに使う。
       ★ W3: source_book は cmd_run が翻訳より前に正規化済み（--dry のときは book と同じ・
       正規化していない）。ここでは正規化をやり直さない。"""
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

    before = None if a.dry else snapshot(source_book)

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
        # ★ 止血3: FREEFORM 経路は no-op ガード/advisories も snapshot() 頼みなので、
        #   検証自体が先頭1000行までしか見ていない（DSL経路より弱い正直さ）。
        notice = _truncation_notice(before, after, exhaustive_postcondition=False)
        if notice:
            print(notice)
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
        # ★ 止血3: 呼び出し元(cmd_run_plan)は lines をそのまま「  {ln}」で表示するだけ
        #   なので、切り詰め注記はここで lines に混ぜて渡す。
        notice = _truncation_notice(before, after, exhaustive_postcondition=False)
        if notice:
            lines = lines + [notice]
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
       （✓ 適用され文書が変化した級に留める＝warn 表示の固定文言で担保）。
       ★ 止血1: 'warn' は2種類の由来を持つ — 語彙外(FREEFORM)段は detail=None の
       固定文言、DSL 段の事後条件が「検証対象が少なすぎる」場合は detail に理由が
       入るのでそちらを見せる（どちらも『機械検証済み』とは言わない点は共通）。"""
    lines = []
    for idx, label, status, detail in items:
        mark = _ITEM_STATUS_MARK[status]
        if status == "ok":
            suffix = f"（{detail}）" if detail else ""
            lines.append(f"{idx}. {label} → {mark} 機械検証済み{suffix}")
        elif status == "warn":
            if detail:
                lines.append(f"{idx}. {label} → {mark} 機械検証できませんでした: {detail}")
            else:
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
        # ★ 止血1: warn の由来は2種類（語彙外の自由生成／DSL段の検証対象不足）ある
        #   ため、どちらにも当てはまる言い方にする。
        return "⚠ 一部は確認が必要です（語彙外の自由生成、または検証対象不足の段があり、機械検証はしていません）", "warn"
    return "✓ すべて機械検証済み", "ok"


def cmd_run_plan(a: argparse.Namespace, book: Path, source_book: Path, book_meta: dict,
                  plan: list) -> int:
    """M2c: 複合依頼の計画実行本体。段ごとに②検証→③確認→④codegen→⑤適用→⑥事後条件
       （DSL 語彙の段）または FREEFORM（語彙外の段・その段の依頼文だけを渡す）を順に実行し、
       ★ 項目別の honest な報告を出す。総合判定は最弱の段に従う（cmd_run_plan 直上の
       overall_verdict）。
       ★ 依存つき連鎖: 各段の接地(verify_dsl_args)は直前までの段を実際に適用した後の
       out_book を読み直した列構成(current_meta)で行う。列名が一致しない場合は
       _apply_new_column_fallback が『直前段が作った新規列』への参照とみなして1回だけ
       書き換えを試みる。
       ★ W3: source_book は cmd_run が翻訳より前に正規化済み（--dry のときは book のまま）。
       header_rows（見出し行の1起点位置）は計画全体を通して不変（並べ替え等は見出し行その
       ものを動かさない）なので、途中の current_meta 再読み込みでも同じ header_rows を渡す。"""
    print(f"■ ailine（複合計画・{len(plan)} 段）  model={a.model}  book={book.name}")

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    out_book = book.with_name(book.stem + ".out" + book.suffix)
    apply_timeout = a.timeout if a.timeout else None
    helpers_dir = Path(a.helpers).resolve() if a.helpers else DEFAULT_HELPERS
    refs_dir = Path(a.refs).resolve() if a.refs else DEFAULT_REFS
    _helper_catalog, helper_files = load_helpers(helpers_dir)
    header_rows = book_meta.get("header_rows", {})
    use_formula = not getattr(a, "values", False)
    vocab = load_vocab()

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
                ok_v, resolved, inferred, err = verify_dsl_args(
                    op, step.get("args", {}), book_meta, task=a.task, vocab=vocab)
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

    shutil.copy2(source_book, out_book)

    original_headers = {k: list(v) for k, v in book_meta["headers"].items()}
    first_sheet = book_meta["sheets"][0] if book_meta.get("sheets") else None
    before_all = snapshot(out_book)
    before_charts = before_all["charts"]

    current_meta = book_meta
    items: list = []         # (idx, label, status, detail)
    plan_json: list = []     # --json 用（既存キー不変・新規追加）
    plan_provenance: list = []   # ★ A': 段ごとの倍率等の出典（history.jsonl 用）

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
            current_meta = build_book_meta(out_book, header_rows=header_rows)
            continue

        # 依存つき連鎖: 直前までの段の適用後の実列構成(current_meta)で接地する
        new_cols = []
        if first_sheet:
            new_cols = [c for c in current_meta["headers"].get(first_sheet, [])
                        if c not in original_headers.get(first_sheet, [])]
        raw_args = step.get("args", {})
        ok_v, resolved, inferred, err = verify_dsl_args(
            op, raw_args, current_meta, task=a.task, vocab=vocab)
        if not ok_v and new_cols and first_sheet:
            patched = _apply_new_column_fallback(
                op, raw_args, current_meta["headers"].get(first_sheet, []), new_cols)
            if patched != raw_args:
                ok_v2, resolved2, inferred2, err2 = verify_dsl_args(
                    op, patched, current_meta, task=a.task, vocab=vocab)
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
        for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
            print(f"  {i}段目: ⚠ {w}")
        if resolved.get("_sources"):
            plan_provenance.append({"step": i, **resolved["_sources"]})
        step_header_row = current_meta.get("header_rows", {}).get(first_sheet, 1) if first_sheet else 1
        code = codegen_dsl(op, resolved, current_meta, use_formula=use_formula)
        (workdir / f"plan_step{i}.bas").write_text(code, encoding="utf-8")

        t0 = progress_start(f"⏳ {i}段目 LibreOffice で適用中…")
        okrun, err_apply, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not okrun:
            detail = f"実行時エラー: {short_error_summary(err_apply)}"
            items.append((i, label, "fail", detail))
            plan_json.append({"op": op, "command": line, "status": "fail", "postcondition": None})
            continue

        status, reason = run_postcondition(op, out_book, resolved, before_charts=before_charts,
                                            header_row=step_header_row, use_formula=use_formula)
        # ★ 止血1/2: status ∈ {"pass","warn","fail","error"}。"error" はチェッカー内の
        #   例外を捕まえた印（段の報告上は fail 扱い・生トレースバックは出さない）。
        if status in ("fail", "error"):
            items.append((i, label, "fail", reason))
            plan_json.append({"op": op, "command": line, "status": "fail", "postcondition": "fail"})
            continue
        if status == "warn":
            # ★ 止血1: 検証対象が少なすぎて意味を持たない → 段の成功は名乗るが
            #   『機械検証済み』とは言わない（format_plan_report の warn+detail 分岐）。
            items.append((i, label, "warn", reason))
            plan_json.append({"op": op, "command": line, "status": "warn", "postcondition": "warn"})
            current_meta = build_book_meta(out_book, header_rows=header_rows)
            continue

        items.append((i, label, "ok", reason))
        plan_json.append({"op": op, "command": line, "status": "ok", "postcondition": "pass"})
        current_meta = build_book_meta(out_book, header_rows=header_rows)

    print()
    for ln in format_plan_report(items):
        print(ln)
    verdict_line, verdict = overall_verdict(items)
    print(f"\n{verdict_line}")

    after_all = snapshot(out_book)
    _changed, difflines = diff_snapshots(before_all, after_all)
    result["plan"] = plan_json
    result["provenance"] = plan_provenance or None
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
    r.add_argument("--values", action="store_true",
                   help="COMPUTE_COLUMN を式でなく値ベタ書きにする（既定は式・W3 Part3）")
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

    v = sub.add_parser("vocab", help="用語集（税率等の取り決め値）を編集・表示する")
    vsub = v.add_subparsers(dest="vocab_cmd", required=True)
    va = vsub.add_parser("add", help="語を登録する（例: ailine vocab add 消費税 1.1）")
    va.add_argument("term", help="語（例: 消費税）")
    va.add_argument("value", help="値（倍率。例: 1.1）")
    vl = vsub.add_parser("list", help="登録済みの語を一覧表示する")
    v.set_defaults(func=cmd_vocab)
    return ap


def main(argv=None) -> int:
    a = build_parser().parse_args(argv)
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())
