#!/usr/bin/env python
"""ailine — 自然言語のタスクを、ローカル LLM が LibreOffice Basic に書き起こし、
   basrun で文書に適用し、★ 効果を読み戻して検証する（「走った ≠ できた」）。

    ailine run  <book> "<タスク>"          生成 → 検品ゲート → 原本に反映 → 変化検証 → 修復 → 差分表示
    ailine run  <book> "<タスク>" --dry     生成して見せるだけ（適用しない・レビュー用）
    ailine run  <book> "<タスク>" --copy    原本には触らず <book>.out に結果を作る（旧既定）
    ailine undo <book>                      直前の反映を取り消す（世代バックアップから復元）
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
- ★ **既定=原本にそのまま反映（W8b-2）**: `<book>` に直接書く。「壊さない」の担保は
  もう「コピーにしか書かない」ことではなく、3重の安全網でまかなう:
  ①往復忠実度ゲート（LO 往復だけで失われる飾りを検品して申告・`--accept-loss`/`--copy` で選ぶ）
  ②世代バックアップ+`ailine undo`（反映前に必ず退避・いつでも戻せる）
  ③ no-op ガード＋事後条件などの機械検証（変化・正しさを見る）。
  従来の「コピーにしか書かない」挙動は `--copy` で選べる（原本は無変更・`<book>.out.xlsx` に生成）。
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
import hashlib
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

# ★ W8b 項目6: グローバル run ロック。基盤の LibreOffice(basrun 経由)が単一インスタンス
#   (port 2002)前提のため、ブック単位でなく `ailine run` 全体で1本にする。
RUN_LOCK_FILE = HISTORY_DIR / "run.lock"
RUN_LOCK_STALE_SECONDS = 30 * 60   # 30分超のロックは stale とみなして奪取する

# ★ W10a 項目2: 既定変更(W8b-2・原本直接反映)の一度きり告知。marker ファイルの有無だけで
#   判定する（history 等には依存しない・単純に「見せたか」の1ビット）。
NOTICE_V2_FILE = HISTORY_DIR / "notice_v2_shown"
NOTICE_V2_TEXT = (
    "★ このバージョンから、既定で原本に直接反映します（自動バックアップ+undo つき）。"
    "従来のコピー方式は --copy。この告知は一度だけ表示されます"
)

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
        "★ W10b: 依頼に無い操作のヘルパは絶対に呼ぶな（『何か効くかもしれない』で総当たりに"
        "色々呼ぶのは禁止。依頼を達成するのに要る分だけを呼ぶ）。\n"
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


# ★ W8a 項目3: 旧文言は「答えて」と聞くだけで答え方(CLI で何を打てばいいか)が無い行き止まり
#   だった（architect 発見）。--header-row フラグの使い方まで添えて、次のコマンドが打てる形にする。
CLARIFY_HEADER_ROW_QUESTION = ("見出しが何行目か分かりません。"
                                "`--header-row 3` のように指定して再実行してください")


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


# ★ W10b 項目4a(摩擦): 「★疑わしい: 変更が元データの範囲外」が、新規列の作成という
#   *意図した*操作でも毎回出ていた実測所見への対応。DSL 経路(COMPUTE_COLUMN・target
#   無指定=新規列作成)に限り、detect_ghost_data が指す範囲が丸ごとその新規列に収まって
#   いれば中立表示に落とす（新規列は原本の使用範囲外に出るのが当然＝誤警報）。
# ★ W10d: COMPUTE_COLUMN 専用の if だったものを OP_WRITE_TARGET の宣言駆動へ一般化した
#   （査定で名指しされたオオカミ少年: LOOKUP_FILL が新規列を作っても同じ誤警報が出ていた）。
#   op ごとの if は増やさない — OP_WRITE_TARGET に既に登録済みの「書き込み先列」宣言を
#   そのまま読み、その列が対象シートの既存見出しに無ければ『新規列を作る効果』とみなす
#   （COMPUTE_COLUMN の target 無指定＝キーが空、LOOKUP_FILL の target_col 有指定だが
#   対象シートにまだ無い列名＝どちらも実行時にその名前で新しい列を作るという同じ意味）。
_GHOST_RANGE_RE = re.compile(r"^★ 疑わしい: 変更が元データの範囲外です（([A-Z]+)\d+(?::([A-Z]+)\d+)?）$")


def _declared_new_column_letter(op: str, resolved: dict, book_meta: dict) -> str | None:
    """op が今回、宣言済みの効果として新規列を作るなら、その列の文字（"C" 等）を返す。
       作らない/対象シートが分からない場合は None。
       ★ W10d 番人の土台: OP_WRITE_TARGET だけを見る。新しい op を足しても
       OP_WRITE_TARGET へ登録さえすれば、ここへの追記なしで正しく判定される
       （test_op_write_target_declares_all_ops が登録漏れ自体を防ぐ）。"""
    write_target = OP_WRITE_TARGET.get(op)
    if not write_target:
        return None
    col_key, sheet_key = write_target
    if sheet_key:
        sheet = resolved.get(sheet_key)
    else:
        sheets = book_meta.get("sheets") or []
        sheet = sheets[0] if sheets else None
    if not sheet:
        return None
    headers = book_meta.get("headers", {}).get(sheet, [])
    col_name = resolved.get(col_key)
    if col_name and col_name in headers:
        return None   # 既存列への書き込み（上書き側の話・新規列ではない）
    new_col_idx = len(headers)   # 0起点・新規列は既存見出しの直後
    return get_column_letter(new_col_idx + 1)


def _neutralize_new_column_ghost_warning(advisories: list, op: str, resolved: dict,
                                          book_meta: dict) -> list:
    """op が今回、宣言済みの効果として新規列1本を作る場合、advisories 中の
       『変更が元データの範囲外』行のうち、範囲が丸ごとその新規列1本に収まっている
       ものだけを中立表示に置き換える。範囲が他列にも及ぶ場合（保守的）や
       op に新規列作成の宣言が無い場合は一切変えない。"""
    new_col_letter = _declared_new_column_letter(op, resolved, book_meta)
    if not new_col_letter:
        return advisories
    out = []
    for line in advisories:
        m = _GHOST_RANGE_RE.match(line)
        if m and m.group(1) == new_col_letter and (m.group(2) or m.group(1)) == new_col_letter:
            out.append("（新規列の追加は意図どおりです）")
            continue
        out.append(line)
    return out


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
       を添える（りんご欠落型のような『1行だけ抜けている』変更を1秒で見えるように）。
       ★ W8a 項目2: 分子(M)は必ずデータ行のみに数える。_used_range/_data_row_count は
       先頭行(min_r)を見出し行とみなして分母を数えているのに、分子側(changed_rows)は
       全変更行をそのまま数えていたため、見出し行も対象列で変わった場合（例: 新規列を
       作って見出し+データ全行に書く COMPUTE_COLUMN）に『データ5行のうち6行を変更』の
       ような算数が壊れた表示になっていた（実測: e2e_work/w3_e2e3_log.txt）。
       見出し行(min_r)自体の変更・原本の使用範囲より下の新規行の変更は、分子(データ行数)
       には数えず、見出し行の変更だけ別語「＋見出し行」で添える。"""
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
    min_r, max_r, min_c, _max_c = rect
    key_col = col - 1 if col > min_c else col
    data_rows = _data_row_count(before, sheet, key_col)
    all_changed_rows = {r for _, r, _ in changed}
    header_changed = min_r in all_changed_rows
    # ★ 分子=データ行のみ: 見出し行(min_r)と、原本の使用範囲より下の新規行(> max_r)は
    #   ここでは「データ行」として数えない（分母(data_rows)も同じ範囲の定義に揃える）。
    data_changed_rows = {r for r in all_changed_rows if min_r < r <= max_r}
    changed_rows = len(data_changed_rows)
    unchanged_rows = max(data_rows - changed_rows, 0)
    col_letter = get_column_letter(col)
    msg = (f"列 {col_letter}: データ {data_rows} 行のうち {changed_rows} 行を変更"
           f"（{unchanged_rows} 行は未変更）")
    if header_changed:
        msg += "＋見出し行"
    return msg


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


# ★ W10c 中: AGGREGATE(SummaryTable)/PIVOT(DataPilot) は定義上・毎回新規シートを作る
#   のが op の宣言済みの効果。依頼文が「シート」「ピボット」「別に」のどれも使わない
#   言い方（例:「部門ごとに金額をまとめて」）だと _NEW_SHEET_MENTION_RE の言及ベース抑制は
#   効かず、意図した新設のたびに「★ 依頼にない新しいシートが作成されました」が出ていた
#   （査定で名指しされた摩擦・W10b 項目4a で COMPUTE_COLUMN の新規列にやったのと同じ処置）。
OP_DECLARED_SHEET_EFFECT = {"AGGREGATE", "PIVOT"}


def _neutralize_declared_new_sheet_warning(advisories: list, op: str, before: dict, after: dict) -> list:
    """op が OP_DECLARED_SHEET_EFFECT（新規シート作成が宣言済みの効果）で、かつ実際に
       ちょうど1枚だけ新規シートができた場合に限り、その1枚についての
       『依頼にない新しいシートが作成されました』を中立表示に落とす。
       ★ 保守的（安全器官の減衰は迷ったら出す側）: 新規シートが2枚以上できた場合や
       op が対象外の場合は一切変えない（宣言どおりの効果と断定できないケースは残す）。"""
    if op not in OP_DECLARED_SHEET_EFFECT:
        return advisories
    new_sheets = _new_sheets(before, after)
    if len(new_sheets) != 1:
        return advisories
    sheet = new_sheets[0]
    target_line = f"★ 依頼にない新しいシートが作成されました（{sheet}）"
    out = []
    for line in advisories:
        if line == target_line:
            out.append(f"（新規シート『{sheet}』の作成は意図どおりです）")
            continue
        out.append(line)
    return out


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


def mention_overlap_advisory(mentions: dict, before: dict, after: dict,
                              exclude_sheets: set | None = None) -> list:
    """言及があるのに変更範囲と全く重ならない場合だけ警告する（保守的）。
       数字表記の列は 0 起点/1 起点の両解釈を許し、どちらかが触られていれば沈黙する。
       ★ W10b 項目4b(摩擦): exclude_sheets に載るシート（例: LOOKUP_FILL の参照専用
       source_sheet）は、依頼文に言及があっても『変更されていません』を出さない
       （読み取り専用が正しい操作で、変更が無いのが正常なため誤警報だった）。
       ★ 安全器官の減衰なので保守的に: 抑制は呼び出し側が op から明示的に渡す時だけ・
       既定(None)は従来どおり無抑制。"""
    exclude_sheets = exclude_sheets or set()
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
        if sheet in exclude_sheets:
            continue
        if sheet not in changed_sheets:
            lines.append(f"★ 依頼で言及された『{sheet}』は存在しません/変更されていません")
    return lines


def _structural_advisories(before: dict, after: dict) -> list:
    """助言のうち『この差分そのものが疑わしいか』を判定する部分だけ
       （①幽霊データ ②一様埋め ③件数の突き合わせ ⑤新規シートの中身・★ W6）。
       依頼文言との重なり(④ mention_overlap_advisory)は含めない。
       ★ W10d: 複合計画(cmd_run_plan)が段ごとの before/after にこの部分だけを
       再利用するために build_advisories から切り出した。④は依頼文全体に対する
       充足を問う質問なので、段ごとの局所的な before/after では判定できない
       （他段が担当する言及まで『この段で変更されていない』と誤検知する）。
       単発 op(build_advisories 経由)ではこれまでどおり④も同じ before/after で
       評価する（そちらは1段しかないため局所=全体で一致し、挙動は不変）。"""
    lines = []
    for fn in (detect_ghost_data, detect_uniform_fill):
        msg = fn(before, after)
        if msg:
            lines.append(msg)
    recon = count_reconciliation(before, after)
    if recon:
        lines.append(recon)
    lines.extend(new_sheet_advisories(before, after))
    return lines


def build_advisories(task: str, before: dict, after: dict, exclude_sheets: set | None = None) -> list:
    """diff の後に表示する助言行を全部集める。
       ①幽霊データ ②一様埋め ③件数の突き合わせ ⑤新規シートの中身（★ W6・
       _structural_advisories が担当） ⑥依頼にないシート新設の申告（★ W6）
       ④依頼文言との重なり。"""
    lines = list(_structural_advisories(before, after))
    lines.extend(unrequested_new_sheet_advisory(task, before, after))
    mentions = extract_task_mentions(task, before["sheets"])
    lines.extend(mention_overlap_advisory(mentions, before, after, exclude_sheets))
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
# ★ W10b 項目3: 「1.1を掛けた/掛ける」「1.1で割った/割る」型（税込み/税抜きの言い換えで
#   「倍」を伴わない場合の実測ギャップ・battery v5 #503 で発覚）。掛けるは n そのまま、
#   割るは 1/n（税抜き＝税込み金額から逆算する倍率）。
_RATE_KAKE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:を|に)?\s*掛け")
_RATE_WARI_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:で|に)?\s*割っ")
_RATE_KEYWORD_RE = re.compile(r"税|倍率")
_RATE_BARE_NUM_RE = re.compile(r"(\d+(?:\.\d+)?)")
# ★ W10c 高: 依頼文に「率らしい語」が一切無いのに COMPUTE_COLUMN の「1列×率」パターン
#   （税込み/税抜き専用）へ誤分類された時に分類そのものを疑うための、より広い信号語。
#   上の各 _RATE_*_RE より緩い（数値を伴わなくても良い）— 「倍率を求める話かどうか」
#   だけを見る。
_RATE_SIGNAL_RE = re.compile(r"[%％]|倍|掛け|割っ|税")
# ★ W10c 中: 新規列の見出しを自然な日本語にするための語（A' 原則: LLM を使わず正規表現の
#   有無判定だけで決める。査定で名指しされた「金額*1.1」という数式風の見出しの対応）。
_TAX_INCLUSIVE_RE = re.compile(r"税込")
_TAX_EXCLUSIVE_RE = re.compile(r"税抜")


def extract_rate_factor(text: str) -> tuple:
    """依頼文から明示の倍率を抽出する。戻り値は (factor, 出典スニペット) か (None, None)。
       ①「10%」「8 ％」型 → 1+n/100 ②「1.1倍」型 → n そのまま ②'「1.1を掛けた」型 → n
       そのまま ②''「1.1で割った」型 → 1/n（税抜き等の逆算） ③「税」「倍率」という語の
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
    for m in _RATE_KAKE_RE.finditer(text):
        f = round(float(m.group(1)), 6)
        candidates.setdefault(f, m.group(0))
    for m in _RATE_WARI_RE.finditer(text):
        n = float(m.group(1))
        if n > 0:
            f = round(1 / n, 6)
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

# ★ W9: op メタデータ（接地フォーム設計の確定分・今回は宣言だけで UI は変えない）。
#   各 op に category(大分類・事務の言葉)・label(操作名・事務の言葉)・synonyms(依頼文で
#   よく使う言い方の例・宣言のみ＝翻訳プロンプトへ機械的に注入はしない) を一箇所で持つ。
#   OP_LABELS は後方互換のためこの dict から導出する（既存コードの OP_LABELS.get(op, op)
#   はそのまま・値は完全に同じ）。
OP_META = {
    "SORT": {"category": "並べ替える", "label": "並べ替え",
              "synonyms": ["並べ替え", "ソート", "順に並べる"]},
    "COMPUTE_COLUMN": {"category": "計算する", "label": "計算列",
                         "synonyms": ["計算", "掛け算・割り算", "列同士の演算"]},
    "LOOKUP_FILL": {"category": "表を編集する", "label": "転記",
                     "synonyms": ["引っ張ってくる", "転記", "VLOOKUP"]},
    "AGGREGATE": {"category": "計算する", "label": "集計",
                   "synonyms": ["集計", "まとめる", "グループごとに小計"]},
    "BOLD": {"category": "見た目を整える", "label": "太字",
              "synonyms": ["太字", "ボールド", "強調"]},
    "FILL_COLOR": {"category": "見た目を整える", "label": "背景色",
                    "synonyms": ["色を付ける", "塗りつぶす", "ハイライト"]},
    "NUMBER_FORMAT": {"category": "見た目を整える", "label": "数値書式",
                        "synonyms": ["桁区切り", "カンマ区切り", "3桁区切り"]},
    "MERGE": {"category": "表を編集する", "label": "セル結合",
               "synonyms": ["結合", "セルを繋げる", "セルをまとめる"]},
    "CHART": {"category": "グラフを作る", "label": "グラフ",
               "synonyms": ["グラフ", "棒グラフ", "チャート"]},
    "CENTER_ALIGN": {"category": "見た目を整える", "label": "中央揃え",
                       "synonyms": ["中央揃え", "センタリング", "真ん中に寄せる"]},
    "APPEND_TOTAL": {"category": "計算する", "label": "合計追加",
                       "synonyms": ["合計を出す", "税込み合計", "一番下に合計"]},
    # ★ W9: 検証済みヘルパ4種の DSL 語彙昇格。
    "INSERT_ROWS": {"category": "表を編集する", "label": "行挿入",
                      "synonyms": ["行を挿入", "行を追加", "行を足す"]},
    "DRAW_BORDERS": {"category": "見た目を整える", "label": "けい線",
                       "synonyms": ["けい線を引く", "罫線を引く", "枠線を付ける"]},
    "AUTOFIT": {"category": "見た目を整える", "label": "列幅自動調整",
                 "synonyms": ["幅を内容に合わせる", "列幅調整", "列を自動調整"]},
    "PIVOT": {"category": "計算する", "label": "ピボット",
               "synonyms": ["ピボットテーブル", "ピボットで集計", "クロス集計"]},
}

OP_LABELS = {op: meta["label"] for op, meta in OP_META.items()}

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
    # ★ W9: INSERT_ROWS は count が既定値ありの任意項目（at だけ必須）。
    #   DRAW_BORDERS/AUTOFIT は引数無し（空タプル＝ any(...) が常に False で FREEFORM に落ちない）。
    "INSERT_ROWS": ("at",),
    "DRAW_BORDERS": (),
    "AUTOFIT": (),
    "PIVOT": ("group_col", "value_col"),
}

# ★ W10c 致命1: 「破壊の関所」（既存列への上書き検知・下の _maybe_warn_target_overwrite）が
#   守る対象を op ごとの if 分岐で持たず宣言駆動にする。旧実装は
#   `if op != "COMPUTE_COLUMN": return None` の1行で、COMPUTE_COLUMN 以外（LOOKUP_FILL 等）
#   は関所が構造的に発火しなかった（監査実測: 存在しない転記先列が無関係な既存列へ
#   解決され、確認なしで上書きされた事故）。
#   値は (書き込み先列の resolved args キー, 対象シート名の resolved args キー or None)、
#   または「この op には既存列の値を上書きする効果が無い」ことを示す明示の None。
#   None は「安全だから省略した」のではなく「対象が無いと確認した」宣言 — 新しい op を
#   足すたびにここへの追記が必須になる（test_op_write_target_declares_all_ops が
#   OP_SCHEMA の全 op に対する宣言漏れを機械的に検査する＝再発防止の本体）。
#   sheet_key が None のときは book_meta の先頭シート（現行 DSL が書き込み対象にする
#   唯一のシート）を指す。LOOKUP_FILL だけ target_sheet で別シートを明示できる。
OP_WRITE_TARGET = {
    "SORT": None,                          # 並べ替えのみ・値そのものは保存される
    "COMPUTE_COLUMN": ("target", None),    # target 無指定時は新規列（resolved に無い＝安全)
    "LOOKUP_FILL": ("target_col", "target_sheet"),
    "AGGREGATE": None,                     # 新規シートを作るだけ（既存列は書かない）
    "BOLD": None,                          # 書式のみ・値を書かない
    "FILL_COLOR": None,
    "NUMBER_FORMAT": None,
    "MERGE": None,
    "CHART": None,
    "CENTER_ALIGN": None,
    "APPEND_TOTAL": None,                  # データ末尾の新規行に追記するだけ（W6・既存列は不可侵）
    "INSERT_ROWS": None,                   # 行を挿入するだけ・既存値は下にずれるだけで残る
    "DRAW_BORDERS": None,
    "AUTOFIT": None,
    "PIVOT": None,                         # 新規シートを作るだけ
}

# ★ bench/translation_spike.py（実測 v1）と同じ語彙定義（bench 側は比較用に据え置き、
#   本番プロンプトはここが唯一の元）。
OPS_DOC = """SORT: 並べ替え。args: col(列名), order(asc|desc)
COMPUTE_COLUMN: 既存列同士の計算。args: operands(列名2つ), operator(+,-,*,/), target(省略可・実在する列名。
  依頼が「〜に」のように既存列を名指ししたらその列名を入れる。無指定なら新しい列を作る)
  ★ 税込み/税抜き等「1列 × 率」の場合は operands を既存列名1つだけの配列にする。
  operator は * (税込み等、掛ける) か / (税抜き等、割る)。倍率(税率等)の数値はここに入れない
  （数値化はここでは行わない・機械が別途確定する。APPEND_TOTAL の倍率と同じ扱い）
LOOKUP_FILL: 別シートの対応表から値を転記。args: target_sheet, target_col, source_sheet, key_col
  ★ target_col は target_sheet に実在する列名が望ましいが、無ければ依頼文に書かれている
  そのままの列名を入れてよい（実行時にその名前で新しい列を作る）。実在しない列名を
  依頼文に無い別の実在列名に置き換えて誤魔化してはいけない（無関係な列を上書きする事故になる）
AGGREGATE: グループ別に集計表を作る。args: group_col, value_col
BOLD: 太字。args: target("row:行番号" か "col:列名")
FILL_COLOR: 背景色。args: target("row:N"か"col:列名"), color(英語色名)
NUMBER_FORMAT: 数値書式。args: col(列名), style("thousands")
MERGE: セル結合。args: range("A1:C1"形式)
CHART: 棒グラフ。args: value_col(列名)
CENTER_ALIGN: 中央揃え。args: target("all" か "col:列名")
APPEND_TOTAL: 列の合計(SUM)を表の最終行の下に追加する（税込み合計等）。args: col(合計する列名),
  label(省略可・既定"合計"。表示ラベル。「税込み合計」等、依頼の言い方をそのまま入れる)
  ★ 合計(SUM)専用。平均・最大・最小など他の統計量は語彙に無い（OUT_OF_VOCAB にする）。
  倍率(税率等)は入れない。数値化はここでは行わない（機械が別途確定する）
INSERT_ROWS: 行を挿入する。依頼文に具体的な行番号がある時だけ使う。
  args: at(1起点の行番号。この行の位置に挿入され、既存行は下にずれる), count(省略可・既定1・挿入する行数)
DRAW_BORDERS: 依頼文に「けい線/罫線/枠線」という言葉が明示された時だけ使う。表にけい線(格子線)を
  引く。args不要（表全体が対象）★「整えて」「いい感じに」のような具体性の無い依頼には
  絶対に使わない（曖昧なら CLARIFY で確認する）
AUTOFIT: 依頼文に「列幅/幅」という言葉が明示された時だけ使う。列幅を内容に合わせて自動調整する。
  args不要（全列が対象）★ 具体性の無い依頼には使わない（曖昧なら CLARIFY で確認する）
PIVOT: 依頼文に「ピボット」「ピボットテーブル」という言葉が明示された時だけ使う。
  本物のピボットテーブル(DataPilot)を作る。args: group_col(分類する列), value_col(合計する列)
  ★「ピボット」の語が無いグループ別集計（「集計」「まとめる」「小計」「合計がみたい」等、
  ピボットという語を伴わない言い方はすべて）は AGGREGATE を使う（PIVOT は LibreOffice で
  開き直すたび書式が消える癖があるため、書式つきの見栄えが要るなら AGGREGATE の方が適する）"""

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
    # ★ W10c 致命2: target_col が対象シートにまだ無いケース（監査実測: これを教えないと
    #   LLM が依頼に無い別の実在列名（この例なら「数量」）を勝手に代入することがあった）。
    ('対象ブックの構成: {"明細": ["商品コード", "数量"], "単価表": ["商品コード", "単価"]}\n'
     '依頼: 「単価表を見て単価を入れて」',
     '{"plan": [{"op": "LOOKUP_FILL", "args": {"target_sheet": "明細", "target_col": "単価", '
     '"source_sheet": "単価表", "key_col": "商品コード"}}]}'),
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
    # ★ W9: 検証済みヘルパ4種の語彙昇格。PIVOT/AGGREGATE の分岐（「ピボット」と明示
    #   された時だけ PIVOT・それ以外の集計語は AGGREGATE）と INSERT_ROWS を1例ずつ教える。
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「部門ごとにピボットテーブルで集計して」',
     '{"plan": [{"op": "PIVOT", "args": {"group_col": "部門", "value_col": "金額"}}]}'),
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「部門ごとに金額をまとめて」',
     '{"plan": [{"op": "AGGREGATE", "args": {"group_col": "部門", "value_col": "金額"}}]}'),
    ('対象ブックの構成: {"Sheet": ["商品", "金額"]}\n'
     '依頼: 「3行目の前に1行挿入して」',
     '{"plan": [{"op": "INSERT_ROWS", "args": {"at": 3, "count": 1}}]}'),
    # ★ W9 回帰対応（battery v1 再実測で実測）: DRAW_BORDERS/AUTOFIT が「引数不要」なぶん
    #   曖昧な依頼の当てはめ先として誤って選ばれやすかった（「整えて」「いい感じにして」→
    #   誤って DRAW_BORDERS を選ぶ誤断定が発生）。CLARIFY が正しい例を明示する。
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「整えて」',
     '{"plan": [{"op": "CLARIFY", "question": "「整える」とは具体的に何をしますか"'
     '"（例: けい線を引く／列幅を合わせる／太字にする）"}]}'),
    # ★ W9 回帰対応: 「ピボット」の語が無いのに漠然とした集計依頼を PIVOT と誤断定した
    #   実測（#14「部署別の売上合計がみたい」→ AGGREGATE が正しい）。同型の言い回しを直接教える。
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「部門別の金額合計がみたい」',
     '{"plan": [{"op": "AGGREGATE", "args": {"group_col": "部門", "value_col": "金額"}}]}'),
    # ★ W9 回帰対応: APPEND_TOTAL は合計(SUM)専用。平均等の他統計量は語彙外
    #   （実測: 「平均値を追加」を誤って APPEND_TOTAL にした）。
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「金額の平均値を一番下に追加して」',
     '{"plan": [{"op": "OUT_OF_VOCAB", "about": "平均値の追加"}]}'),
    # ★ W10b 項目3: 税込み/税抜き等「1列 × 率」パターンの語彙昇格（operator 第6回査定の
    #   実測: この言い回しが4種すべて DSL 分類に失敗し自由生成へ退避していた）。
    #   few-shot は最小限（W9 の実測: 足しすぎると別 op の誤断定回帰が出る）。
    #   ★ A': 倍率の数値化は LLM に求めない（factor は machine-determined。verify_dsl_args の
    #   extract_rate_factor/lookup_vocab_factor が依頼文/用語集から機械確定する）。
    #   ★ battery v3 再走で実測した回帰: 「税込み○○を出して」型の言い回しで例を作ると、
    #   APPEND_TOTAL の label 例（「税込み合計を一番下に出して」）と表現が似すぎて、
    #   モデルが APPEND_TOTAL の label フィールドまで省略するようになった(#305 CLARIFY
    #   すべきが確定に化けた)。「〜を追加して」型（新規列作成が明確な言い回し）の例だけ
    #   残し、APPEND_TOTAL の言い回しと重ならないようにする。
    ('対象ブックの構成: {"Sheet": ["部門", "金額"]}\n'
     '依頼: 「金額に1.1を掛けた列を追加して」',
     '{"plan": [{"op": "COMPUTE_COLUMN", "args": {"operands": ["金額"], "operator": "*"}}]}'),
    # ★ W10c 高: 「列を全部Xに書き換える」（数値の倍率ではなく文字列を一律に代入する）は、
    #   税込み/税抜き（COMPUTE_COLUMN の1列×率パターン）と表現が似ているが別物＝語彙に無い
    #   （実測: 「氏名の列を全部『退職済み』に書き換えて」が税率の話と誤認された事故）。
    ('対象ブックの構成: {"Sheet": ["氏名", "部署", "金額"]}\n'
     '依頼: 「氏名の列を全部『退職済み』に書き換えて」',
     '{"plan": [{"op": "OUT_OF_VOCAB", "about": "列を同じ文字列で一括上書き"}]}'),
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
- 「整えて」「いい感じにして」のような、何をすればよいか具体的に書いていない依頼は、
  引数が要らない操作(DRAW_BORDERS/AUTOFIT 等)であっても絶対に推測で当てはめない。
  必ず CLARIFY で何をしたいか確認する
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
        # ★ W10b 項目3: 税込み/税抜き等「1列 × 率」パターン。operands が列名1つだけの
        #   配列なら『既存列×倍率』とみなす（2列の四則演算とは別モード）。倍率(factor)は
        #   APPEND_TOTAL と同じ A' 原則で LLM から受け取らず機械確定する
        #   （extract_rate_factor/lookup_vocab_factor・regex のみ）。
        single_factor_mode = isinstance(operands, list) and len(operands) == 1
        if not single_factor_mode and not (isinstance(operands, list) and len(operands) == 2):
            return False, resolved, inferred, "演算対象が2つの列名になっていません"

        if single_factor_mode:
            v, was_inferred, err = resolve_col_ref(operands[0], headers.get(first_sheet, []))
            if err:
                return False, resolved, inferred, err
            resolved["operands"] = [v]
            if was_inferred:
                inferred.add("operands")
            if resolved.get("operator") not in ("*", "/"):
                return False, resolved, inferred, (
                    f"演算子『{resolved.get('operator')}』は列1つの計算（税込み/税抜き等）"
                    "では * か / のみ対応です")

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
                # ★ W10c 高: 依頼文に率らしい語が一切無いのに「1列×率」（税込み/税抜き専用）へ
                #   分類されているのは、分類そのものが誤っている可能性が高い（実測: 「氏名の
                #   列を全部『退職済み』に書き換えて」のような値の一括書き換え依頼が、税率の
                #   話と誤認されて COMPUTE_COLUMN の単列モードに落ちることがあった）。
                #   その場合は「倍率が分からない」でなく、分類自体を疑う文言に変える
                #   （率を要求する op に分類されたのに率の手がかりが無い＝CLARIFY の理由を
                #   正直に言い換える。指示は意図・保証は機械＝プロンプト側だけに頼らない）。
                if not _RATE_SIGNAL_RE.search(task or ""):
                    return False, resolved, inferred, (
                        f"依頼「{task}」は『{v}』列に何らかの倍率（税率等）を掛ける操作として"
                        "解釈しましたが、依頼文に倍率らしき手がかりが見当たりません。"
                        "列の値をそのまま書き換える操作は今のところ対応していません。"
                        "倍率を掛ける処理であれば、依頼文に率を書く（例:「消費税10%」）か、"
                        "用語集に登録してください（例: ailine vocab add 消費税 1.1）"
                    )
                return False, resolved, inferred, (
                    "倍率（税率等）が分かりません。依頼文に率を書く（例:「消費税10%」）か、"
                    "用語集に登録してください（例: ailine vocab add 消費税 1.1）"
                )
            if resolved["factor"] <= 0:
                return False, resolved, inferred, f"倍率『{resolved['factor']}』は正の数でなければなりません"
            if sources:
                resolved["_sources"] = sources
            # ★ W10c 中: 新規列の見出しの自然化。旧実装は見出しを f"{op1}{operator}{factor:g}"
            #   （例:「金額*1.1」）という数式風の文字列にしていた（査定で名指し）。target
            #   無指定(新規列作成)かつ依頼文が税込み/税抜きと分かる言い方の場合だけ、
            #   その日本語ラベルを見出しに使う（A' 原則: LLM を使わず正規表現の有無のみで
            #   決める。手がかりが無ければ従来どおりの数式風見出しにフォールバック）。
            if not resolved.get("target"):
                if _TAX_INCLUSIVE_RE.search(task or ""):
                    resolved["_new_col_label"] = f"税込{v}"
                elif _TAX_EXCLUSIVE_RE.search(task or ""):
                    resolved["_new_col_label"] = f"税抜{v}"
                elif _RATE_KEYWORD_RE.search(task or ""):
                    resolved["_new_col_label"] = f"税込{v}" if resolved["operator"] == "*" else f"税抜{v}"
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
            raw_target = resolved["target"]
            v, was_inferred, err = resolve_col_ref(raw_target, headers.get(first_sheet, []))
            if err:
                if "一意に決まりません" in err:
                    return False, resolved, inferred, err
                del resolved["target"]
            else:
                resolved["target"] = v
                if was_inferred:
                    inferred.add("target")
                    # ★ W10a 項目3: 数字指定→列名解決の元の表記を残す（解釈要約の表示用・
                    #   例:「列5」→「在庫」列と解決した時、確認行の直後にその経緯を見せる）。
                    resolved["_target_raw"] = raw_target

    elif op == "LOOKUP_FILL":
        if (err := check_sheet("target_sheet")):
            return False, resolved, inferred, err
        if (err := check_sheet("source_sheet")):
            return False, resolved, inferred, err
        if resolved["target_sheet"] != first_sheet:
            return False, resolved, inferred, f"対象シートは1枚目（{first_sheet}）のみ対応しています"
        # ★ W10c 致命2: target_col は COMPUTE_COLUMN の target と違い OP_SCHEMA 上は必須
        #   slot なので、LLM は「存在しないなら空にする」を選べない。実測（監査再現）:
        #   対象シートに『単価』列がまだ無いのに転記を頼むと、LLM がそれと無関係な
        #   *実在する*既存列（例:「数量」）の名前を代わりに返すことがある。resolve_col_ref
        #   は実在列名なら無条件で素通しするため、これだけでは見分けられない（そのまま
        #   進めると「数量」が確認なしで上書きされる事故になる）。
        #   ここでは「実在するから信用する」をやめ、根拠を要求する:
        #   ①依頼文にその列名が書かれている ②転記元（source_sheet）の値列
        #   （VLookupFromTable ヘルパの仕様どおり常に列1＝2番目の列）と同じ名前
        #   のどちらかが無いと、実在列であっても信用しない。
        target_headers = headers.get(resolved["target_sheet"], [])
        source_headers = headers.get(resolved["source_sheet"], [])
        value_col_hint = source_headers[1] if len(source_headers) > 1 else None
        raw_target_col = resolved.get("target_col")
        raw_str = str(raw_target_col) if raw_target_col not in (None, "") else ""
        exists = raw_str in target_headers
        mentioned = bool(raw_str) and raw_str in task
        matches_value_col = value_col_hint is not None and raw_str == value_col_hint

        if exists and (mentioned or matches_value_col):
            pass   # 根拠つきで実在列を指名＝そのまま使う（上書き注意は破壊の関所が別途担当）
        elif not exists:
            cands = _digit_candidates(raw_str, target_headers)
            if len(cands) == 1:
                resolved["target_col"] = cands[0]   # 数字表記の推定は従来どおり許容
                inferred.add("target_col")
            elif mentioned:
                # ★ 依頼文にも同じ列名が書かれている＝新規作成が正しい解釈（COMPUTE_COLUMN の
                #   target 無指定＝新規列と同じ考え方）。target_col はそのまま残し、
                #   codegen_dsl 側で新規列として作る。
                pass
            else:
                known = ", ".join(target_headers) if target_headers else "(無し)"
                return False, resolved, inferred, (
                    f"転記先の列『{raw_target_col}』が『{resolved['target_sheet']}』シートに"
                    f"見つかりません。ある列: {known}。新しい列として作る場合は、依頼文に"
                    f"その列名を書いてください（例:「{raw_target_col}という列を作って転記して」）"
                )
        else:
            # ★ 実測の事故そのもの: exists=True だが根拠が無い（依頼文にも書かれておらず、
            #   転記元の値列とも一致しない）＝上書き対象を取り違えている可能性が高い。
            hint = f"（参照表『{resolved['source_sheet']}』の値の列は『{value_col_hint}』です）" \
                if value_col_hint else ""
            return False, resolved, inferred, (
                f"転記先の列『{raw_target_col}』は実在しますが、依頼文にその列名が見当たらず、"
                f"転記元の値とも対応が確認できません{hint}。上書き対象を取り違えている"
                "可能性があるため、意図した列名を依頼文に明記してください"
            )
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

    # --- ★ W9: 検証済みヘルパ4種の語彙昇格 -----------------------------------
    elif op == "INSERT_ROWS":
        at_raw = str(resolved.get("at", "")).strip()
        if not (at_raw.isdigit() and int(at_raw) >= 1):
            return False, resolved, inferred, f"行番号『{resolved.get('at')}』が不正です（1以上の整数）"
        resolved["at"] = int(at_raw)
        count_raw = resolved.get("count")
        if count_raw in (None, ""):
            resolved["count"] = 1
            inferred.add("count")
        else:
            count_str = str(count_raw).strip()
            if not (count_str.isdigit() and int(count_str) >= 1):
                return False, resolved, inferred, f"挿入行数『{count_raw}』が不正です（1以上の整数）"
            resolved["count"] = int(count_str)

    elif op == "DRAW_BORDERS":
        pass   # 引数無し・表全体が対象（検証することが無い）

    elif op == "AUTOFIT":
        pass   # 引数無し・全列が対象（検証することが無い）

    elif op == "PIVOT":
        # ★ AGGREGATE と同じ2 slot（group_col/value_col）。列名の実在確認だけ共有する。
        if (err := resolve_in("group_col", first_sheet)):
            return False, resolved, inferred, err
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
    # ★ W8a 項目5: 表示ラベルのみ「倍率」→「率」（税率・掛け率の文脈での事務向け言い換え）。
    #   内部キー("factor")・関数名・コメントは不変。
    "APPEND_TOTAL": (("対象列", "col", None), ("ラベル", "label", None), ("率", "factor", None)),
    # ★ W9: 検証済みヘルパ4種の語彙昇格。
    "INSERT_ROWS": (("挿入位置", "at", None), ("行数", "count", None)),
    "DRAW_BORDERS": (),
    "AUTOFIT": (),
    "PIVOT": (("分類列", "group_col", None), ("集計列", "value_col", None)),
}

# ★ W9 項目4: PIVOT(DataPilot) の既知の癖（README 記載・再描画で書式が撥ねる）を
#   確認行・結果表示の両方に一言添える。AGGREGATE(SummaryTable) との使い分けを促す。
PIVOT_CAVEAT = "書式なしの素の表になります。書式つきは『集計表』"


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
        src = resolved_args["source_sheet"].replace('"', '""')
        target_col_name = resolved_args["target_col"]
        # ★ W10c 致命2: target_col が対象シートに実在しない場合（verify_dsl_args が依頼文に
        #   同じ列名があると確認済み）は、COMPUTE_COLUMN の新規列作成と同じ考え方で
        #   末尾に新しい列を作ってから転記する（無関係な既存列を上書きしない）。
        if target_col_name in theaders:
            tgt_idx = theaders.index(target_col_name)
            header_write = ""
        else:
            tgt_idx = len(theaders)   # 0起点・次の空き列
            header_name = str(target_col_name).replace('"', '""')
            header_write = (f'    oDoc.Sheets.getByIndex(0).getCellByPosition({tgt_idx}, {hr0})'
                             f'.setString("{header_name}")\n')
        return _wrap_basic(header_write +
                            f'    Call VLookupFromTable(oDoc, {hr0}, {key_idx}, {tgt_idx}, "{src}")\n')

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
        operands = resolved_args["operands"]
        operator = resolved_args["operator"]
        target = resolved_args.get("target")

        if len(operands) == 1:
            # ★ W10b 項目3: 税込み/税抜き等「列 × 率」パターン（1列 + 機械確定した factor）。
            #   codegen は既存の2列版と同じ書き方（式=セル参照、値ベタ書きの選択も同じ）。
            op1 = operands[0]
            i1 = headers[first_sheet].index(op1)
            factor = float(resolved_args["factor"])
            if target:
                new_col = headers[first_sheet].index(target)
                header_write = ""
            else:
                new_col = len(headers[first_sheet])
                # ★ W10c 中: verify_dsl_args が税込み/税抜きと判定できていれば自然な
                #   日本語見出し（例:「税込金額」）を使う。無ければ従来どおりの数式風見出し。
                header_name = str(resolved_args.get("_new_col_label")
                                   or f"{op1}{operator}{factor:g}").replace('"', '""')
                header_write = f"    oSheet.getCellByPosition({new_col}, {hr0}).setString(\"{header_name}\")\n"
            if use_formula:
                col1_letter = get_column_letter(i1 + 1)
                write_line = (f'        oSheet.getCellByPosition({new_col}, i).setFormula('
                              f'"=" & "{col1_letter}" & (i + 1) & "{operator}{factor:g}")\n')
            else:
                write_line = (f"        oSheet.getCellByPosition({new_col}, i).setValue("
                              f"oSheet.getCellByPosition({i1}, i).getValue() {operator} {factor:g})\n")
            body = ("    Dim oSheet As Object, lastRow As Long, i As Long\n"
                    "    oSheet = oDoc.Sheets.getByIndex(0)\n"
                    + _scan_last_row_basic(start_row=str(hr0 + 1))
                    + header_write
                    + f"    For i = {hr0 + 1} To lastRow\n"
                    + write_line
                    + "    Next i\n")
            return _wrap_basic(body)

        op1, op2 = operands
        i1 = headers[first_sheet].index(op1)
        i2 = headers[first_sheet].index(op2)
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

    # --- ★ W9: 検証済みヘルパ4種の語彙昇格。いずれも helpers/*.bas のヘルパは headerRow
    #   引数を取らない（物理1行目を前提に自前走査する既存実装・ここでは変更しない）ため、
    #   codegen 側も hr0 を渡さずそのまま Call するだけ。
    if op == "INSERT_ROWS":
        at0 = int(resolved_args["at"]) - 1   # 1起点(Excel行番号) → 0起点(Basic)
        count = int(resolved_args.get("count", 1) or 1)
        return _wrap_basic(f"    Call InsertRows(oDoc, {at0}, {count})\n")

    if op == "DRAW_BORDERS":
        return _wrap_basic("    Call DrawTableBorders(oDoc)\n")

    if op == "AUTOFIT":
        return _wrap_basic("    Call AutoFitColumns(oDoc)\n")

    if op == "PIVOT":
        g_idx = headers[first_sheet].index(resolved_args["group_col"])
        v_idx = headers[first_sheet].index(resolved_args["value_col"])
        return _wrap_basic(f"    Call PivotSum(oDoc, {g_idx}, {v_idx})\n")

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
       無いケースの両方を見逃さない）。
       ★ W10b 項目3: operands が1つだけ（税込み/税抜き等「列 × 率」パターン）の場合は
       check_compute_column_single_factor に委譲する。"""
    if len(args["operands"]) == 1:
        return check_compute_column_single_factor(path, args, header_row=header_row,
                                                    use_formula=use_formula)
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


def check_compute_column_single_factor(path: Path, args: dict, header_row: int = 1,
                                        use_formula: bool = False) -> tuple:
    """★ W10b 項目3: COMPUTE_COLUMN の「1列 × 率」パターン（税込み/税抜き等）専用の事後条件。
       check_compute_column（2列版）と同じ二層検証（式の期待形・data_only キャッシュ値）を
       1列 + factor に合わせて行う。factor は verify_dsl_args が機械確定済みの値
       （resolved["factor"]）をそのまま受け取る前提。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    op1 = args["operands"][0]
    operator = args["operator"]
    factor = float(args.get("factor", 1) or 1)
    i1 = _col_index_by_header(ws, op1, header_row=header_row)
    target = args.get("target")
    # ★ W10c 中: codegen_dsl と同じ見出し名決定（_new_col_label があれば自然な日本語見出し）。
    newname = target or args.get("_new_col_label") or f"{op1}{operator}{factor:g}"
    inew = _col_index_by_header(ws, newname, header_row=header_row)
    if i1 is None or inew is None:
        wb.close()
        return "fail", f"演算対象または対象列『{newname}』が見つからない"
    last = _scan_last_row(ws, header_row=header_row)
    wb_v = None
    ws_v = None
    if use_formula:
        wb_v = openpyxl.load_workbook(path, data_only=True)
        ws_v = wb_v[wb_v.sheetnames[0]]
    col1_letter = get_column_letter(i1)
    checked = 0
    excluded = 0
    for r in range(header_row + 1, last + 1):
        a = ws.cell(row=r, column=i1).value
        got = ws.cell(row=r, column=inew).value
        if not _is_number(a):
            excluded += 1   # 例: 合計行で演算対象セルが空欄
            continue
        want = _apply_operator(a, factor, operator)
        if use_formula:
            expect_formula = f"={col1_letter}{r}{operator}{factor:g}"
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
        # ★ W10b 項目4a: 摩擦(operator 実測) — マスタ表の列順が値→キーだと1件も引けず
        #   ここに落ちるが、旧メッセージは原因を示さなかった。VLookupFromTable ヘルパは
        #   常に「参照表は列0=キー・列1=値」固定（列順非依存化は他列数のマスタで誤ヒット
        #   するリスクが高くリスク高と判断・具体誘導での対応を選んだ＝W10b 報告参照）。
        return "fail", (
            f"対応表『{args['source_sheet']}』に載っているキーが1件も転記されていません。"
            f"マスタ表はキー列→値列の順である必要があります。『{args['source_sheet']}』"
            f"シートの A 列にキー（{args['key_col']} に対応する値）、B 列に値を置いてください"
        )
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


# --- ★ W9: 検証済みヘルパ4種の事後条件 ---------------------------------------
#   ヘルパ本体(helpers/*.bas)が headerRow を取らないのと同じ理由で、これらのチェッカーは
#   openpyxl の生スキャン(_scan_last_row/_scan_last_col)で実データ範囲を都度見つける
#   （header_row を渡しはするが、ヘルパが物理1行目前提で動く以上、通常は 1 のまま使う想定）。

def check_insert_rows(path: Path, args: dict, header_row: int = 1,
                       source_book: Path | None = None) -> tuple:
    """INSERT_ROWS の事後条件。source_book(適用前のコピー)が渡されれば、
       ①データ最終行が count 増えている ②挿入位置(at)以降の内容が、適用前の対応行から
       count 行分ずれて現れている（シフト）③挿入された行自体が空欄、を突き合わせる。
       ★ source_book が無い（複合計画の途中段等で before が用意できない経路）場合は、
       挿入位置が空欄であることだけを見る warn 判定に落とす（保守的・断定しない）。"""
    at = int(args["at"])
    count = int(args.get("count", 1) or 1)

    if source_book is None or not Path(source_book).exists():
        wb = openpyxl.load_workbook(path)
        ws = wb[wb.sheetnames[0]]
        last_col = max(_scan_last_col(ws, header_row=header_row), 1)
        row_cells = [ws.cell(row=at, column=c).value for c in range(1, last_col + 1)]
        wb.close()
        if all(v in (None, "") for v in row_cells):
            return "warn", "挿入位置が空欄であることのみ確認（適用前ファイルとの突き合わせ無し）"
        return "fail", f"{at}行目が空欄でない（挿入されていない可能性）"

    wb_before = openpyxl.load_workbook(source_book)
    ws_before = wb_before[wb_before.sheetnames[0]]
    last_before = _scan_last_row(ws_before, header_row=header_row)
    last_col = _scan_last_col(ws_before, header_row=header_row)
    if last_col < 1 or last_before < header_row + 1:
        wb_before.close()
        return "fail", _ZERO_TARGET_REASON

    wb_after = openpyxl.load_workbook(path)
    ws_after = wb_after[wb_after.sheetnames[0]]

    # ★ 挿入は AFTER 側に意図的な空行（挿入行そのもの）を作るため、連続データを前提にする
    #   _scan_last_row を AFTER 側の「最終行」検出には使わない（挿入行で即座に打ち切られて
    #   しまう）。BEFORE の各データ行が「count 行下」に正しく現れているかを直接照合する。
    mismatches = checked = 0
    for r in range(max(at, header_row + 1), last_before + 1):
        for c in range(1, last_col + 1):
            checked += 1
            if ws_before.cell(row=r, column=c).value != ws_after.cell(row=r + count, column=c).value:
                mismatches += 1
    inserted_ok = all(ws_after.cell(row=r, column=c).value in (None, "")
                       for r in range(at, at + count) for c in range(1, last_col + 1))
    # 期待される最終データ行のさらに1行下が空欄であること（想定外の余剰データが無いこと）。
    expect_last_after = last_before + count
    extra_row_empty = all(ws_after.cell(row=expect_last_after + 1, column=c).value in (None, "")
                           for c in range(1, last_col + 1))
    wb_before.close(); wb_after.close()

    if checked == 0:
        return "fail", _ZERO_TARGET_REASON
    if mismatches:
        return "fail", f"挿入位置以降のシフトが一致しない（不一致 {mismatches}/{checked} セル）"
    if not inserted_ok:
        return "fail", f"挿入された{count}行が空欄でない"
    if not extra_row_empty:
        return "fail", f"期待より多くの行にデータがある（{expect_last_after}行より下に想定外のデータ）"
    return "pass", f"{count}行挿入・{checked}セル分のシフトを確認（挿入行は空欄・行数も+{count}ぴったり）"


def check_draw_borders(path: Path, args: dict, header_row: int = 1) -> tuple:
    """DRAW_BORDERS の事後条件。使用範囲の全セルに上下左右の罫線属性が付いているかを見る
       （DrawTableBorders ヘルパは格子罫線を範囲全体に一括で付けるため、1セルでも
       欠けていれば ヘルパが動いていないか範囲がずれている）。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    last_row = _scan_last_row(ws, header_row=header_row)
    last_col = _scan_last_col(ws, header_row=header_row)
    if last_col < 1 or last_row < header_row:
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    cells = [ws.cell(row=r, column=c) for r in range(header_row, last_row + 1)
             for c in range(1, last_col + 1)]
    wb.close()
    if not cells:
        return "fail", _ZERO_TARGET_REASON

    def _has_border(cell) -> bool:
        bd = cell.border
        return bool(bd and bd.top and bd.top.style and bd.bottom and bd.bottom.style
                     and bd.left and bd.left.style and bd.right and bd.right.style)

    ok = all(_has_border(c) for c in cells)
    if not ok:
        return "fail", "使用範囲の一部に罫線が無い"
    return "pass", f"{len(cells)} セルの罫線を確認"


def check_autofit(path: Path, args: dict, header_row: int = 1,
                   source_book: Path | None = None) -> tuple:
    """AUTOFIT の事後条件。source_book が渡されれば、使用中の各列の幅が適用前後で
       変化した列数を数える（1列でも変化していれば pass）。source_book が無ければ、
       AFTER 側で幅が明示的に設定されている列があることだけを見る warn 判定に落とす。"""
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    last_col = _scan_last_col(ws, header_row=header_row)
    if last_col < 1:
        wb.close()
        return "fail", _ZERO_TARGET_REASON
    after_widths = {}
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        after_widths[letter] = dim.width if dim and dim.width else None
    wb.close()

    if source_book is not None and Path(source_book).exists():
        wb_b = openpyxl.load_workbook(source_book)
        ws_b = wb_b[wb_b.sheetnames[0]]
        before_widths = {}
        for c in range(1, last_col + 1):
            letter = get_column_letter(c)
            dim = ws_b.column_dimensions.get(letter)
            before_widths[letter] = dim.width if dim and dim.width else None
        wb_b.close()
        changed = sum(1 for k in after_widths if after_widths[k] != before_widths.get(k))
        if changed == 0:
            return "fail", "列幅が変化していない"
        return "pass", f"{changed}/{len(after_widths)} 列の幅が変化"

    set_count = sum(1 for v in after_widths.values() if v)
    if set_count == 0:
        return "fail", "列幅が設定されていない（AutoFitColumns が効いていない可能性）"
    return "warn", f"{set_count} 列に幅が設定されていることのみ確認（適用前ファイルとの突き合わせ無し）"


def check_pivot(path: Path, args: dict, header_row: int = 1) -> tuple:
    """PIVOT の事後条件。①『ピボット』シートが存在 ②xlsx zip 内に本物の DataPilot
       (xl/pivotTables/) が実在するか、を見る（W8b の忠実度ゲートと同じ zip 直読み技法・
       LO を再起動しない）。★ W9 項目4: DataPilot は開き直すたび再描画で値キャッシュが
       変わりうる既知の癖があるため、集計値そのものの照合（check_aggregate 相当）はせず
       構造の存在確認に留める（PIVOT_CAVEAT として案内も添える）。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    has_sheet = "ピボット" in wb.sheetnames
    wb.close()
    if not has_sheet:
        return "fail", "『ピボット』シートが無い"
    try:
        with zipfile.ZipFile(path) as z:
            names = [n.lower() for n in z.namelist()]
    except Exception as e:
        return "fail", f"xlsx を zip として読めない: {e}"
    has_pivot_table = any(n.startswith("xl/pivottables/") for n in names)
    if not has_pivot_table:
        return "fail", "『ピボット』シートはあるが DataPilot(ピボットテーブル)の実体が無い"
    return "pass", f"『ピボット』シートと DataPilot を確認（{PIVOT_CAVEAT}）"


POSTCONDITIONS = {
    "SORT": check_sort, "COMPUTE_COLUMN": check_compute_column,
    "LOOKUP_FILL": check_lookup_fill, "AGGREGATE": check_aggregate,
    "BOLD": check_bold, "FILL_COLOR": check_fill_color,
    "NUMBER_FORMAT": check_number_format, "MERGE": check_merge,
    "CENTER_ALIGN": check_center_align, "APPEND_TOTAL": check_append_total,
    # ★ W9: 検証済みヘルパ4種。
    "INSERT_ROWS": check_insert_rows, "DRAW_BORDERS": check_draw_borders,
    "AUTOFIT": check_autofit, "PIVOT": check_pivot,
}


def run_postcondition(op: str, out_book: Path, resolved_args: dict, before_charts: int = 0,
                       header_row: int = 1, use_formula: bool = False,
                       source_book: Path | None = None) -> tuple:
    """⑥ op 別事後条件。(status, reason)。status ∈ {"pass","warn","fail","error"}。
       CHART だけ before_charts と比較する専用の形。
       ★ W3: header_row（1起点、省略時1）を全チェッカーに一貫して渡す（『三層全部が
       同じ見出し推定を使う』の事後条件側）。use_formula は COMPUTE_COLUMN 専用（W3 Part3）。
       ★ W9: source_book（適用前のコピー・無ければ None）は INSERT_ROWS/AUTOFIT だけが使う
       （before/after の突き合わせが要る2op・他は無視される安全なキーワード引数）。
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
        if op in ("INSERT_ROWS", "AUTOFIT"):
            return fn(out_book, resolved_args, header_row, source_book=source_book)
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

# ★ W8b: 秒精度(旧・6桁)・マイクロ秒精度(新・12桁)・衝突回避の "-N" 連番つき、の
#   いずれも受け付ける（後方互換）。
_BACKUP_TS_RE = re.compile(r"^\d{8}T\d{6}(?:\d{6})?Z(?:-\d+)?$")
_BACKUP_TS_SEQ_RE = re.compile(r"-(\d+)$")


def _utc_ts() -> str:
    """ファイル名に使える UTC タイムスタンプ（例: 20260814T120000123456Z）。
       ★ W8b: マイクロ秒まで含める（実測: それでも Windows の壁時計分解能は粗く
       20万回の連続呼び出しで56通りしか値が変わらない。衝突自体は make_backup 側の
       "-N" 連番フォールバックで最終的に防ぐ・ここは表示上の精度向上に留める）。"""
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%f") + "Z"


def _ts_sort_key(ts: str):
    """ts 文字列（秒/マイクロ秒精度・衝突回避の "-N" 連番つき）を新しい順ソート用の
       比較可能な値 (datetime, 連番) にする。壊れた形式は最古扱い。"""
    body = ts
    seq = 0
    m = _BACKUP_TS_SEQ_RE.search(body)
    if m:
        seq = int(m.group(1))
        body = body[:m.start()]
    if body.endswith("Z"):
        body = body[:-1]
    for fmt in ("%Y%m%dT%H%M%S%f", "%Y%m%dT%H%M%S"):
        try:
            return (datetime.strptime(body, fmt), seq)
        except ValueError:
            continue
    return (datetime.min, seq)


def _backup_namespace(book: Path) -> str:
    """book が置かれているフォルダごとのバックアップ名前空間（sha1 の先頭8桁）。
       ★ W8b 項目3: 同名ファイルが別フォルダにある場合（A\\見積.xlsx / B\\見積.xlsx）、
       旧・フラットな BACKUP_DIR では stem+suffix だけで一致させていたため取り違え
       （undo 混線）が起きえた。フォルダの絶対パスでハッシュ化した名前空間ディレクトリへ
       分離する（book が相対パスで渡されても resolve() してから使う）。"""
    return hashlib.sha1(str(Path(book).resolve().parent).encode("utf-8")).hexdigest()[:8]


def backup_path_for(book: Path, ts: str | None = None) -> Path:
    ts = ts or _utc_ts()
    return BACKUP_DIR / _backup_namespace(book) / f"{book.stem}.{ts}{book.suffix}"


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
    """book のバックアップを ~/.ailine/backups/<名前空間>/ に作る。戻り値はバックアップ先。
       ★ 失敗したら例外を投げる（呼び出し側が --inplace 中止の判断に使う）。
       ★ M2c: 新しいバックアップを作った後、keep 世代を超えた古いものを剪定する
       （既定 DEFAULT_KEEP_BACKUPS=10。無制限にすると個人開発機のディスクを静かに食う）。
       ★ W8b 項目3: 新規のバックアップは必ず名前空間ディレクトリへ書く（フラット領域には
       もう書かない・読み取り専用の後方互換は list_backups 側で担う）。
       ★ W8b: Windows の壁時計分解能は実測で粗く（20万回の連続呼び出しで56通りしか
       値が変わらない）、restore_backup が「復元前の現状」を退避する高速な連続呼び出し
       等でファイル名が衝突しうる。衝突したら "-N" 連番を足して必ず別ファイルにする
       （既存の世代を上書きで消さない・回帰テストで自己顕在化した実際の不具合の修正）。"""
    ts = _utc_ts()
    dst = backup_path_for(book, ts=ts)
    dst.parent.mkdir(parents=True, exist_ok=True)
    n = 2
    while dst.exists():
        dst = backup_path_for(book, ts=f"{ts}-{n}")
        n += 1
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
    """book に対応するバックアップを新しい順(タイムスタンプ降順)で返す。
       ★ W8b 項目3: 名前空間ディレクトリ BACKUP_DIR/<ns>/ を主として見る。
       旧フラット領域（BACKUP_DIR 直下・名前空間分離前の名残）も読み取り専用互換で
       あわせて見る（新規はもう書かない・iterdir で拾うのはファイルだけ＝名前空間の
       サブディレクトリ自体を誤ってバックアップと数えないよう is_file() で絞る）。"""
    stem, suffix = book.stem, book.suffix
    found = []
    ns_dir = BACKUP_DIR / _backup_namespace(book)
    if ns_dir.is_dir():
        for p in ns_dir.iterdir():
            ts = _parse_backup_name(p.name, stem, suffix)
            if ts is not None:
                found.append((ts, p))
    if BACKUP_DIR.is_dir():
        for p in BACKUP_DIR.iterdir():
            if not p.is_file():
                continue
            ts = _parse_backup_name(p.name, stem, suffix)
            if ts is not None:
                found.append((ts, p))
    # ★ W8b: 秒精度(旧)とマイクロ秒精度(新)が混在しうるため、生文字列の辞書順ではなく
    #   _ts_sort_key() でパースした実時刻順に並べる（桁数違いの文字列比較は時刻順にならない）。
    found.sort(key=lambda pair: _ts_sort_key(pair[0]), reverse=True)
    return [p for _ts, p in found]


def restore_backup(book: Path) -> Path:
    """book を「1つ前の世代」から復元する。戻り値は使ったバックアップの Path。
       バックアップが1つも無ければ例外を投げる。
       ★ W8b-2: 連続 undo が実編集履歴を1段ずつ正しく遡れるようにする（B2 事故バッテリ
       ①で実測: 素朴に『常に最新のバックアップを使う』実装だと、undo が作る『復元前の
       現状の退避』が次回の undo で『最新』として選ばれてしまい、2回目の undo が
       1段先(=1回目の undo を打ち消す)に進んでしまっていた＝多段 undo が成立しない）。
       book の現在の中身と一致するバックアップがあれば、そのすぐ内側(より古い方)を
       復元先にする（＝現在地がバックアップ履歴のどこかに『既にいる』とみなし、そこから
       もう1段遡る）。一致するものが無ければ（＝直前の実編集の直後・通常の最初の undo）
       最新のバックアップを復元する。復元前の現状は必ず退避する（undo 自体も可逆）。"""
    backups = list_backups(book)   # 新しい順
    if not backups:
        raise FileNotFoundError(f"{book.name} のバックアップが無い")

    target = backups[0]
    if book.exists():
        try:
            current_bytes = book.read_bytes()
        except OSError:
            current_bytes = None
        if current_bytes is not None:
            for i, p in enumerate(backups):
                try:
                    matched = p.read_bytes() == current_bytes
                except OSError:
                    continue
                if matched:
                    target = backups[i + 1] if i + 1 < len(backups) else backups[i]
                    break
        make_backup(book)   # 復元前の現状も退避＝restore 自体も可逆にする
    shutil.copy2(target, book)
    return target


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


def cmd_undo(a: argparse.Namespace) -> int:
    """★ W8b 項目5: `restore` の昇格。真実の源はバックアップファイル自体
       （history.jsonl には依存しない＝history が壊れていても undo できる）。
       名前空間対応(item3)は list_backups/restore_backup 経由でそのまま効く。
       復元後、まだ戻せる回数（＝現時点で使えるバックアップの総数）を添える。"""
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
    remaining = len(list_backups(book))
    print(f"✓ {book.name} を {used.name} から復元した（あと {remaining} 回戻せます）")
    return 0


# ---------------------------------------------------------------------------
# ★ W8b: 安全器官（既定の反転は次コミット。今回は原本を直接書く危険を減らす下ごしらえ）
# ---------------------------------------------------------------------------

def check_excel_lock(book: Path) -> str | None:
    """book が Excel 等で開かれている兆候を機械的に見る。開かれていそうなら理由の
       文字列（人間可読）、そうでなければ None。
       ★ W8b 項目2: ①同フォルダの Excel ロックファイル(~$<name>) の存在
       ②open(book, 'r+b') を試みて PermissionError になるか、の2つを見る（保守的
       ＝どちらかに該当したら『開かれている可能性』として止める。誤検知より、
       書き込み中の文書を壊さない方を優先する）。run の最初（LO 起動・翻訳より前）
       に呼ぶ（--copy 時も含め常に同じ判定にする＝整合性の観点で経路を分けない）。"""
    lock_file = book.parent / f"~${book.name}"
    if lock_file.exists():
        return f"ロックファイル {lock_file.name} が存在します"
    try:
        with open(book, "r+b"):
            pass
    except PermissionError:
        return "書き込みロックされています（他のアプリが開いている可能性）"
    except OSError:
        return None   # その他の I/O エラーはここでは判定しない（保守的・誤検知回避）
    return None


# --- ★ W8b 項目1: 往復忠実度ゲート -------------------------------------------
#   normalize_book の LO 往復『だけ』で失われる飾り（原本にはあり、正規化後には無い）を
#   検出する。マクロの効果とは無関係（normalize は何もしない空マクロで一度 LO を通す
#   だけの工程）。喪失ゼロなら無言（体験は不変）。喪失があれば --inplace の直前で
#   申告し、--accept-loss / --copy のどちらかを選ばせる（exit 4）。

_FIDELITY_RELS_RE = re.compile(r"^xl/worksheets/_rels/.*\.rels$", re.I)


def _fidelity_zip_members(path: Path) -> set:
    try:
        with zipfile.ZipFile(path) as z:
            return {n.lower() for n in z.namelist()}
    except Exception:
        return set()


def _classify_fidelity_member(name: str) -> str | None:
    """xlsx zip 内の1エントリが『LO 往復で失われがちな飾り』のどのカテゴリか。
       対象外（本文と無関係な構造ファイル等）なら None。"""
    if name.startswith("xl/pivotcache") or name.startswith("xl/pivottables/"):
        return "ピボットテーブル"
    if name.startswith("xl/drawings/"):
        return "図形/描画"
    if name.startswith("xl/media/"):
        return "画像"
    if name == "xl/vbaproject.bin":
        return "VBA マクロ"
    if _FIDELITY_RELS_RE.match(name):
        return "リンク情報(_rels)"
    return None


def check_zip_fidelity_loss(original: Path, normalized: Path) -> list:
    """(a) original の zip 構成要素のうち normalized で消えたものをカテゴリ別に集計する。
       戻り値: [(カテゴリ, 消えた件数), ...]（件数0のカテゴリは含めない・カテゴリ名順）。
       ★ zip として読めない（xlsx でない・壊れている）場合は空リスト（保守的・誤検知しない）。"""
    before = _fidelity_zip_members(original)
    if not before:
        return []
    after = _fidelity_zip_members(normalized)
    lost = before - after
    counts: dict = {}
    for name in lost:
        cat = _classify_fidelity_member(name)
        if cat:
            counts[cat] = counts.get(cat, 0) + 1
    return sorted(counts.items())


def _count_cf_and_dv(wb) -> tuple:
    """ブック全体の(条件付き書式の件数, 入力規則の件数)を合計する。"""
    cf_total = 0
    dv_total = 0
    for name in wb.sheetnames:
        ws = wb[name]
        cf = getattr(ws, "conditional_formatting", None)
        if cf is not None:
            cf_total += sum(1 for _ in cf)
        dv = getattr(ws, "data_validations", None)
        if dv is not None:
            dv_total += len(dv.dataValidation)
    return cf_total, dv_total


def check_openpyxl_fidelity_loss(original: Path, normalized: Path) -> list:
    """(b) openpyxl で条件付き書式/入力規則の件数が正規化後に減っていないかを見る。
       戻り値: [(カテゴリ, before件数, after件数), ...]（減っていないカテゴリは含めない）。
       ★ どちらか一方でも開けない場合は空リスト（保守的・誤検知しない）。"""
    try:
        wb_b = openpyxl.load_workbook(original)
        cf_b, dv_b = _count_cf_and_dv(wb_b)
        wb_b.close()
        wb_a = openpyxl.load_workbook(normalized)
        cf_a, dv_a = _count_cf_and_dv(wb_a)
        wb_a.close()
    except Exception:
        return []
    out = []
    if cf_a < cf_b:
        out.append(("条件付き書式", cf_b, cf_a))
    if dv_a < dv_b:
        out.append(("入力規則", dv_b, dv_a))
    return out


def check_round_trip_fidelity(original: Path, normalized: Path) -> dict:
    """往復忠実度ゲート本体。{"lost": bool, "items": [{"label":str,"count":int}, ...]}。
       ★ history.jsonl の fidelity フィールドにそのまま記録する形（機械可読）。"""
    items = []
    for cat, n in check_zip_fidelity_loss(original, normalized):
        items.append({"label": cat, "count": n})
    for cat, b, a in check_openpyxl_fidelity_loss(original, normalized):
        items.append({"label": cat, "count": b - a, "before": b, "after": a})
    return {"lost": bool(items), "items": items}


def format_fidelity_warning(fidelity: dict) -> str:
    """人間可読の申告文（例:「⚠ このファイルには、処理すると失われる飾りがあります
       （条件付き書式 3 件・図形/描画 1 件）」）。"""
    parts = "・".join(f"{it['label']} {it['count']} 件" for it in fidelity.get("items", []))
    return f"⚠ このファイルには、処理すると失われる飾りがあります（{parts}）"


# --- ★ W8b 項目4: アトミック置換（--inplace の torn-write 窓の根治） --------------

def atomic_replace_inplace(book: Path, out_book: Path, workdir: Path,
                            keep_backups: int = DEFAULT_KEEP_BACKUPS) -> tuple:
    """--inplace の実体。(ok: bool, error_message: str|None)。ok=False の場合、
       out_book はそのまま残り、原本(book)は無変更（呼び出し側が『--inplace は中止』
       として表示する）。
       ★ 手順: ①バックアップ（失敗したら即中止・原本は触らない）
       ②原本と同じボリューム上の staging(workdir/staged<suffix>)へコピー
       ③os.replace(staging, book) で原子的に置換（POSIX rename(2)/Windows
       MoveFileEx の原子性保証＝torn write の窓が無い。同一ボリュームである
       ことは staging を book と同じ親フォルダ配下(workdir)に置くことで保証する）。
       os.replace が失敗したら（バックアップは既に確保済みの上で）shutil.copy2
       へフォールバックし、その旨を1行表示する。成功時は out_book を削除する
       （原本に反映済みの内容と同じものを残しておく理由が無い＝旧 shutil.move
       と同じ最終状態）。"""
    try:
        make_backup(book, keep=keep_backups)
    except Exception as e:
        return False, f"バックアップに失敗したため --inplace を中止した（原本は無変更）: {e}"

    staging = workdir / f"staged{book.suffix}"
    try:
        shutil.copy2(out_book, staging)
        os.replace(staging, book)
    except OSError as e:
        try:
            shutil.copy2(out_book, book)
        except OSError as e2:
            return False, f"置換に失敗した（バックアップは確保済み）: {e2}"
        print(f"⚠ 原子的な置換に失敗したため copy2 へフォールバックした（バックアップは確保済み）: {e}")
    finally:
        if staging.exists():
            try:
                staging.unlink()
            except OSError:
                pass

    try:
        if out_book.exists() and out_book != book:
            out_book.unlink()
    except OSError:
        pass
    return True, None


# --- ★ W8b 項目6: グローバル run ロック --------------------------------------

def _pid_alive(pid: int) -> bool:
    """PID が生きているか（確実な保証は無いが十分・追加の依存(psutil 等)は増やさない）。
       判定できない場合は「生きている」扱い（安全側＝奪取しない）。"""
    if os.name == "nt":
        try:
            out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"],
                                  capture_output=True, text=True, encoding="utf-8", errors="replace")
            return str(pid) in out.stdout
        except Exception:
            return True
    else:
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except OSError:
            return True


def _read_lock_info(path: Path) -> dict | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _lock_is_stale(info: dict | None) -> bool:
    """既存ロックが奪取してよい(stale)か。①壊れた/読めないロック ②pid が自分自身
       （同一プロセス内の前回呼び出しが解放し損ねた・テスト等で頻出）
       ③pid が既に存在しない ④30分超、のいずれか。"""
    if info is None:
        return True
    pid = info.get("pid")
    if not isinstance(pid, int):
        return True
    if pid == os.getpid():
        return True
    if not _pid_alive(pid):
        return True
    try:
        ts = datetime.fromisoformat(info.get("ts", ""))
        age = (datetime.now(timezone.utc) - ts).total_seconds()
        if age > RUN_LOCK_STALE_SECONDS:
            return True
    except Exception:
        pass
    return False


def acquire_run_lock(path: Path | None = None) -> tuple:
    """(acquired: bool, message: str|None)。O_EXCL でグローバル実行ロックを取る。
       ★ W8b 項目6: 基盤の LibreOffice は単一インスタンス(port 2002)前提のため、
       ブック単位でなく ailine run 全体で1本にする。stale（pid が既に無い/自分自身/
       30分超）なら奪取して新規取得する。"""
    p = path or RUN_LOCK_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    info = {"pid": os.getpid(), "ts": datetime.now(timezone.utc).isoformat(timespec="seconds")}

    def _try_create() -> bool:
        try:
            fd = os.open(str(p), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            return False
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(json.dumps(info))
        return True

    if _try_create():
        return True, None

    existing = _read_lock_info(p)
    if not _lock_is_stale(existing):
        pid = existing.get("pid") if existing else "?"
        ts = existing.get("ts") if existing else "?"
        return False, f"別の ailine が実行中です（pid={pid}・{ts}）"

    try:
        p.unlink()
    except OSError:
        pass
    if _try_create():
        return True, None
    return False, "別の ailine が実行中です（ロック取得の競合）"


def release_run_lock(path: Path | None = None) -> None:
    p = path or RUN_LOCK_FILE
    try:
        p.unlink()
    except OSError:
        pass


def maybe_show_notice_v2(path: Path | None = None) -> bool:
    """★ W10a 項目2: 既定変更(原本直接反映)の一度きり告知。marker ファイルが無ければ
       表示して作る（以後は無言）。戻り値は「今回表示したか」（テスト用）。
       marker の読み書き失敗（権限等）で run 本体を落とさないよう、書き込み失敗は無視する
       （その場合は次回また表示されうる＝安全側に倒す）。
       ★ path 省略時は HISTORY_FILE.parent（＝実運用では NOTICE_V2_FILE と同じ
       ~/.ailine/）から求める（固定の NOTICE_V2_FILE を直接使わない）。実測: この関数は
       cmd_run のたび無条件に呼ぶため、HISTORY_FILE を monkeypatch 済みの大量のテストが
       それに便乗して安全になる（このファイルは backups と違って『一度書いたら消えない』
       ため、実ユーザーの ~/.ailine/notice_v2_shown をテストが汚すと『初回表示』が二度と
       出せなくなる＝汚染の被害が他の一時ファイルより重い）。"""
    p = path or (HISTORY_FILE.parent / "notice_v2_shown")
    if p.exists():
        return False
    print(NOTICE_V2_TEXT)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(datetime.now(timezone.utc).isoformat(timespec="seconds"), encoding="utf-8")
    except OSError:
        pass
    return True


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
        # ★ W10a 項目2: 動作可否のチェックではなく設定の告知（常に ok=True）。
        #   既定が原本直接反映に変わったこと(W8b-2〜)を doctor でも確認できるようにする。
        ("既定動作", True, "原本直接（v2〜）。--copy で従来のコピー方式"),
    ]


# ★ W8a 項目5: doctor の内部名（"ollama 到達 (URL)" 等）は operator（事務職）には
#   意味が伝わらない。内部名は残したまま(技術者が追える)、事務向けの一行説明を併記する。
#   prefix 一致（name には URL/モデル名など動的な値が混ざるため完全一致にしない）。
#   ★ ダミー名（テストの "a"/"b" 等）はどれにも一致しないため従来どおり内部名のみ表示する。
_DOCTOR_BUSINESS_NOTES = (
    ("python 3.10+", "実行に必要なプログラム言語が使えます"),
    ("openpyxl", "Excel ファイルを読み書きする部品が使えます"),
    ("ollama 到達", "AI エンジン (ollama) に接続できています"),
    ("モデル", "使う AI モデルの準備ができています"),
    ("LibreOffice", "文書を開いて処理する土台があります"),
    ("basrun.py", "文書に処理を適用する仕組みがあります"),
    ("demo/", "動作確認用のサンプル文書があります"),
)


def _doctor_business_note(name: str) -> str | None:
    for prefix, note in _DOCTOR_BUSINESS_NOTES:
        if name.startswith(prefix):
            return note
    return None


def format_doctor_report(results: list) -> tuple:
    """(表示テキスト, all_ok)。"""
    lines = []
    all_ok = True
    for name, ok, detail in results:
        mark = "✓" if ok else "×"
        note = _doctor_business_note(name)
        shown = f"{name}（{note}）" if note else name
        if ok:
            line = f"{mark} {shown}" + (f" ({detail})" if detail else "")
        else:
            all_ok = False
            line = f"{mark} {shown}" + (f" — {detail}" if detail else "")
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
       ため、詳細を追いたい時の唯一の入り口になる）。
       ★ W8a 項目1: "dry" は --dry（見せるだけ・未適用）で走ったかどうか。result["dry"] を
       そのまま bool 化する（result に無ければ False＝実適用と同じ扱い）。旧 history.jsonl
       の行（このキーが無い）は read_history/format_history_table 側で dict.get("dry", False)
       により実適用扱いのまま読める（後方互換・新旧を混在させても壊れない）。
       ★ W8b 項目1: "fidelity" は往復忠実度ゲートの検出結果（--inplace が要求され、
       かつ --dry でない run でのみ実際に計算される）。ゲートを走らせなかった run は
       None のまま（既存キーと同じ形＝無ければ None）。"""
    return {
        "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "book": str(book),
        "task": task,
        "model": model,
        "ok": bool(result.get("ok")),
        "dry": bool(result.get("dry")),
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
        "fidelity": result.get("fidelity"),
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
    """人が読める表形式。履歴が無ければ「履歴はまだ無い」を返す。
       ★ W8a 項目1: dry(下見・未適用)の行は末尾に「(下見)」を付けて実適用と区別する
       （「dry-run」は事務の言葉ではないため表示には出さない）。dict にキーが無い旧行は
       dict.get("dry", False) で実適用扱いのまま読める（後方互換）。"""
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
        if e.get("dry", False):
            line += "  (下見)"
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
    # ★ W8b 項目1: 往復忠実度ゲートの検出結果（cmd_run 側で計算済みなら a._fidelity に
    #   乗っている）を history に写す。ゲートを走らせなかった run は None のまま。
    if "fidelity" not in result:
        result["fidelity"] = getattr(a, "_fidelity", None)
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


# ★ W8b-2 項目1: 既定=原本直接適用の終端メッセージを一箇所に集約する
#   （cmd_run_dsl/cmd_run_freeform/cmd_run_plan の3箇所が同じ形だったのを統合）。
#   pending/confirm の中間状態は作らない — undo 一本（architect 判定）。
def _finish_apply(a: argparse.Namespace, book: Path, out_book: Path, workdir: Path,
                   result: dict, machine_verified: bool) -> bool:
    """--copy（a.inplace が False）なら .out のまま（原本は無変更）。既定(a.inplace)なら
       backup+原子的置換(atomic_replace_inplace)で原本へ反映する。
       machine_verified=True（DSL/plan・ルールベース）→「✓ 反映しました」、
       False（自由生成・機械保証なし）→「⚠ 反映しましたが機械保証はありません」。
       戻り値: 置換が成功した(または --copy で置換不要だった)か。"""
    if not a.inplace:
        print(f"\n適用先: {out_book.name}（原本 {book.name} は無変更）")
        result["out"] = str(out_book)
        return True

    ok_ip, err_ip = atomic_replace_inplace(
        book, out_book, workdir, keep_backups=getattr(a, "keep_backups", DEFAULT_KEEP_BACKUPS))
    if not ok_ip:
        print(f"× {err_ip}")
        print(f"適用先: {out_book.name}（原本への反映は中止・原本 {book.name} は無変更）")
        result["out"] = str(out_book)
        return False

    if machine_verified:
        print("\n✓ 反映しました（もとに戻す: ailine undo）")
    else:
        print("\n⚠ 反映しましたが機械保証はありません — 確認して、違えば ailine undo")
    result["out"] = str(book)
    return True


def cmd_run(a: argparse.Namespace) -> int:
    """run コマンドの入口。★ W8b 項目6: 実処理(_cmd_run_body)の前後をグローバル run
       ロックで挟む（基盤の LibreOffice が単一インスタンス前提のため、ブック単位でなく
       ailine run 全体で1本）。取得できなければ即 exit 6（LO 起動・翻訳より前）。
       finally で確実に解放する（sys.exit も SystemExit 例外なので finally は通る）。"""
    acquired, msg = acquire_run_lock()
    if not acquired:
        print(f"× {msg}")
        return 6
    try:
        return _cmd_run_body(a)
    finally:
        release_run_lock()


def _cmd_run_body(a: argparse.Namespace) -> int:
    """run コマンドの本体。★ W3: 正規化パス(＋StructDump による見出し行推定)を
       翻訳より前に一度だけ行う（『三層全部が同じ見出し推定を使う』ための土台。
       translate_task 自身の接地(book_meta)も検出した見出し行を使う）。
       --dry は従来どおり LibreOffice に触れない（見出しは物理1行目のまま・E2E 対象外）。
       ★ W8a 項目3: --header-row N（1起点）が指定されていれば StructDump 検出を丸ごと
       スキップしてその行を採用する（CLARIFY の行き止まりから抜ける唯一の入り口）。
       ★ W8b 項目2: Excel ロック検出は LO 起動・翻訳より前（一番最初）に行う。
       ★ W8b 項目1: --inplace が要求され --dry でない場合だけ、正規化直後に往復忠実度
       ゲートを見る（原本にまだ一切触れていない段階＝『原本に触る前に申告と選択』）。
       ★ W8b 項目4: 自分の workdir(.ailine_<stem>/) は run の終了時に必ず掃除する
       （成功・失敗とも・GC の作り込みはしない＝自分の後始末だけ）。
       ① 見出し行推定（StructDump） → 自信不足なら CLARIFY して exit 3
       ② 翻訳（計画）→
       - 計画が空/1段で CLARIFY → 質問して exit 3
       - 計画が空/1段で DSL 語彙 → ③〜⑥の決定論パイプライン(cmd_run_dsl)
       - 計画が空/1段でそれ以外(FREEFORM・翻訳失敗) → 現行の自由生成経路(cmd_run_freeform)
       - 計画が2段以上(複合依頼) → 段ごとに honest な項目別実行(cmd_run_plan)（M2c）
       ★ 後方互換: translate_task が "plan" で包まない旧形式（bare {"op":...}）を返した場合
       （テストの monkeypatch を含む）も、その dict をそのまま単一段として扱う。"""
    maybe_show_notice_v2()   # ★ W10a 項目2: 既定変更の一度きり告知（run の一番最初）

    book = Path(a.book).resolve()
    if not book.exists():
        sys.exit(f"文書が無い: {book}")

    lock_reason = check_excel_lock(book)
    if lock_reason:
        print(f"× Excel で開かれています。閉じてから実行してください。（{lock_reason}）")
        return 5

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    try:
        return _cmd_run_dispatch(a, book, workdir)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def _cmd_run_dispatch(a: argparse.Namespace, book: Path, workdir: Path) -> int:
    # ★ W8b-2 項目2: 既定 = 原本直接適用（旧 --inplace 相当）。--copy で旧 .out 挙動を
    #   温存する。--inplace は廃止した旧フラグだが、互換のため受理はして移行メッセージ
    #   だけ出す（指定してもしなくても挙動は同じ＝既定が原本適用）。
    if getattr(a, "inplace", False):
        print("★ --inplace は廃止されました。既定で原本に直接適用します"
              "（従来の .out 挙動が欲しい場合は --copy を使ってください）。")
    a.inplace = not getattr(a, "copy", False)

    apply_timeout = a.timeout if a.timeout else None   # 0 で無効化（旧挙動 = 無制限）

    source_book = book
    struct_dump: dict = {}
    if not a.dry:
        t0 = progress_start("⏳ 初回準備（文書の正規化+構造読み取り）…")
        source_book = normalize_book(book, workdir, timeout=apply_timeout)
        progress_end(t0)
        struct_dump = build_struct_dump(source_book, workdir)

        # ★ W8b 項目1: 原本に実際に適用する時だけ、まだ触れていないこの時点で往復忠実度
        #   を見る（喪失ゼロなら無言・体験は不変）。★ W8b-2 項目4: --copy 時は原本に
        #   一切触れないためゲート自体を走らせない（無駄な zip/openpyxl 比較も省く）。
        if a.inplace:
            fidelity = check_round_trip_fidelity(book, source_book)
            a._fidelity = fidelity
            if fidelity["lost"]:
                print(format_fidelity_warning(fidelity))
                if getattr(a, "accept_loss", False):
                    print("→ --accept-loss 指定のため続行します（失われても ailine undo で元に戻せます）")
                else:
                    print("この処理を続けるには、以下のいずれかを指定して再実行してください:")
                    print("  --accept-loss  失われてよい（バックアップから ailine undo で復元可能）")
                    print("  --copy         原本には触らず .out に結果を作る（原本は無変更）")
                    return 4

    sheets = build_book_meta(source_book).get("sheets", [])
    forced_header_row = getattr(a, "header_row", None)
    if forced_header_row:
        # ★ W8a 項目3: --header-row 指定時は検出(StructDump ヒューリスティクス)を丸ごと
        #   スキップし、その行を1枚目シートの見出しとして採用する（他シートは既定1行目のまま
        #   ＝ resolve_header_rows の既定と同じ扱い）。CLARIFY には絶対に落ちない。
        header_rows = {s: 1 for s in sheets}
        if sheets:
            header_rows[sheets[0]] = forced_header_row
        clarify_q = None
    else:
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


def _column_existing_value_count(book_path: Path, sheet_name: str, col_name: str,
                                  header_row: int = 1) -> int:
    """★ M2c / W10a: target(既存列指定)列に、見出し行を除いて値が入っているセルの件数。
       上書き検知の明示（確認行の注意書き）と W10a の破壊の関所（確認メッセージの件数）が
       共有する。読めない/列やシートが見つからない場合は 0（保守的に『無い』扱い＝誤って
       止めない）。★ W3: header_row(1起点)で見出しの実位置を受け取る
       （省略時は物理1行目・旧挙動と同一）。"""
    try:
        wb = openpyxl.load_workbook(book_path, read_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return 0
        ws = wb[sheet_name]
        idx = _col_index_by_header(ws, col_name, header_row=header_row)
        if idx is None:
            wb.close()
            return 0
        last = _scan_last_row(ws, header_row=header_row)
        count = sum(1 for r in range(header_row + 1, last + 1)
                    if ws.cell(row=r, column=idx).value not in (None, ""))
        wb.close()
        return count
    except Exception:
        return 0


def _column_has_existing_values(book_path: Path, sheet_name: str, col_name: str,
                                 header_row: int = 1) -> bool:
    """★ M2c: target(既存列指定)列に、見出し行を除いてどれか値が入っているか
       （_column_existing_value_count の bool 版。他コードとの互換のため残す）。"""
    return _column_existing_value_count(book_path, sheet_name, col_name, header_row=header_row) > 0


def _maybe_warn_target_overwrite(op: str, resolved: dict, book_meta: dict, book_path: Path) -> str | None:
    """★ M2c 項目2 / W10c 致命1: OP_WRITE_TARGET が宣言する書き込み先列に既存値がある場合、
       上書きになる旨の1行を返す（無ければ None・確認行に明示するため）。
       ★ W10c: 対象を COMPUTE_COLUMN 専用の if から OP_WRITE_TARGET の宣言読み取りへ
       一般化した（監査実測: LOOKUP_FILL がこの関所を素通りしていた事故の再発防止。
       OP_WRITE_TARGET のコメント参照）。
       ★ W10a 項目1: この検出（と件数）を「破壊の関所」（原本適用時に確認を挟む・
       cmd_run_dsl/cmd_run_plan 側）がそのまま流用する（検出ロジックを二重管理しない）。"""
    write_target = OP_WRITE_TARGET.get(op)
    if not write_target:
        return None
    col_key, sheet_key = write_target
    col_name = resolved.get(col_key)
    if not col_name:
        return None
    if sheet_key:
        sheet_name = resolved.get(sheet_key)
    else:
        sheets = book_meta.get("sheets") or []
        sheet_name = sheets[0] if sheets else None
    if not sheet_name:
        return None
    header_row = book_meta.get("header_rows", {}).get(sheet_name, 1)
    count = _column_existing_value_count(book_path, sheet_name, col_name, header_row=header_row)
    if count > 0:
        return f"★ 対象列『{col_name}』には既存の値が {count} 件あります（上書きします）"
    return None


def _interpretation_summary_line(resolved: dict, inferred: set) -> str | None:
    """★ W10a 項目3: 実行前の解釈要約。target が数字表記から列名へ推定解決され、かつ
       （呼び出し側が別途 _maybe_warn_target_overwrite で）既存データありと分かっている
       場合だけ、その経緯を1文で見せる（「監査要望3」＝数字指定→列名解決の可視化）。
       それ以外（target が最初から実在列名で指定された等）は None（何も語ることが無い）。"""
    if "target" not in inferred or resolved.get("_target_raw") is None:
        return None
    return f"→『{resolved['_target_raw']}』は既存の『{resolved['target']}』列と解釈しました（既存データあり）"


def _confirm_overwrite_or_gate(a: argparse.Namespace, warn_overwrite: str | None,
                                step_prefix: str = "") -> int | None:
    """★ W10a 項目1: 破壊の関所。既定(原本へ直接反映)で、既存データへの上書きが起きる
       操作（_maybe_warn_target_overwrite が検出）は、--ask 無指定でも確認を挟む
       （監査実測: target が誤って既存列に解決され、確認なしで実データが上書きされた
       事故の再発防止）。--copy/--dry 時は原本に触れない/何もしないため素通し。
       --overwrite が既に立っていれば承知の上として素通し。--ask が既に立っている場合は
       cmd_run_dsl 側の汎用確認で兼ねる（二重に聞かない）。
       戻り値: 続行してよければ None、中断すべきなら呼び出し側がそのまま return すべき
       exit code（対話で拒否=1・非対話で確認できない=7）。
       step_prefix は複合計画の段番号表示用（例: "  2段目: "）。単発 DSL では空文字。"""
    if not (warn_overwrite and getattr(a, "inplace", False) and not getattr(a, "dry", False)
            and not getattr(a, "ask", False) and not getattr(a, "overwrite", False)):
        return None
    try:
        ans = input(f"{step_prefix}上書きしますか？ [y/N]: ").strip().lower()
    except EOFError:
        print(f"{step_prefix}この処理を続けるには、以下のいずれかを指定して再実行してください:")
        print(f"{step_prefix}  --overwrite  上書きを承知して続行する（バックアップから ailine undo で戻せる）")
        print(f"{step_prefix}  --copy       原本には触らず .out に結果を作る（原本は無変更）")
        return 7
    if ans not in ("y", "yes"):
        print(f"{step_prefix}× 中止した")
        return 1
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
    if op == "PIVOT":   # ★ W9 項目4: 確認行の直後にも既知の癖を一言添える。
        print(f"（{PIVOT_CAVEAT}）")
    warn_overwrite = _maybe_warn_target_overwrite(op, resolved, book_meta, book)
    if warn_overwrite:
        summary = _interpretation_summary_line(resolved, inferred)   # ★ W10a 項目3
        if summary:
            print(summary)
        print(warn_overwrite)
    for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
        print(f"⚠ {w}")

    gate_exit = _confirm_overwrite_or_gate(a, warn_overwrite)   # ★ W10a 項目1: 破壊の関所
    if gate_exit is not None:
        return gate_exit

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
    # ★ W8a 項目5: 「決定論」はユーザー向け文字列から排除（内部の設計語彙のまま出すと
    #   事務職には伝わらない）。内部名・コメント・関数名（codegen_dsl 等）は不変。
    print(f"\n─ 生成した .bas（ルール変換・LLM不使用）───────────────")
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
    # ★ W10b 項目4b(摩擦): LOOKUP_FILL の参照専用シート(source_sheet)は書き換えない
    #   （読み取り専用が正しい操作）ので「変更されていません」の対象から除外する。
    exclude_sheets = {resolved["source_sheet"]} if op == "LOOKUP_FILL" else None
    advisories = build_advisories(a.task, before, after, exclude_sheets=exclude_sheets)
    # ★ W10b 項目4a(摩擦): COMPUTE_COLUMN の新規列作成は宣言どおりの効果なので、
    #   その新規列1本に収まる『範囲外』警報は中立表示に落とす。
    advisories = _neutralize_new_column_ghost_warning(advisories, op, resolved, book_meta)
    # ★ W10c 中: AGGREGATE/PIVOT の新規シート作成も同じ考え方で中立表示に落とす。
    advisories = _neutralize_declared_new_sheet_warning(advisories, op, before, after)
    for adv in advisories:
        print(adv)
    result["changes"] = lines
    result["advisories"] = advisories

    status, reason = run_postcondition(op, out_book, resolved, before_charts=before["charts"],
                                        header_row=header_row, use_formula=use_formula,
                                        source_book=source_book)
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

    # ★ W8b-2: DSL 経路はルールベース codegen（自由生成ではない）ので、postcondition が
    #   warn(検証対象不足)でも trailing メッセージは常に ✓「反映しました」側を使う
    #   （postcondition の warn/pass は上ですでに正直に出し分けている＝別レイヤ）。
    _finish_apply(a, book, out_book, workdir, result, machine_verified=True)

    _finish_run(a, book, result, "none")
    return 0


# ---------------------------------------------------------------------------
# ★ W8a 項目4: 率リテラルの機械スキャン（判断棚から昇格）
#   自由生成(FREEFORM/OUT_OF_VOCAB・単段)された Basic コードに、依頼文にも用語集にも
#   出典の無い『率らしい数値』（消費税8%のつもりの 0.08、税込計算の *1.1 等）が紛れて
#   いないかを機械で見る。実測: 単段の自由生成では「8% 仮定」がノーチェックで
#   「✓できました」を素通りしていた（勝手な税率仮定はモデルの幻覚と見分けがつかない）。
#   ★ 保守的: コメント行(') は対象外・整数(小数点なし)は対象外。発火してもブロックはせず
#   「検算してください」の助言止まり（no-op ガード等と違い、正しさの断定はしない）。
# ---------------------------------------------------------------------------

_BASIC_COMMENT_LINE_RE = re.compile(r"^\s*(?:'|REM\b)", re.I)
_RATE_LITERAL_RE = re.compile(r"(?<![\w.])(\d+\.\d+)(?![\w.])")

# ★ W10b 項目2: 修復ループが行き詰まると『総当たりで色々試す』方向に暴走し、依頼と無関係な
#   操作（ヘルパの全呼び出し等）を追加することがあった実測（operator 第6回査定・3回連続
#   再現。詳細は _known_helper_names/detect_helper_sweep のコメント）。修復メッセージ自体に、
#   依頼の範囲を超えるな、という制約を明示的に加える（no-op/実行時エラーの2種のみ・
#   bad_signature/truncated は構造の修正だけなので対象外）。
_REPAIR_SCOPE_GUARD = "★ 元の依頼に無い操作(無関係なヘルパ呼び出しを含む)を追加してはいけない。"


def _looks_like_rate(value: float) -> bool:
    """0.05〜0.2（例: 消費税8%→0.08）または 1.05〜1.2（例: 税込計算の *1.1）の小数か。"""
    return (0.05 <= value <= 0.2) or (1.05 <= value <= 1.2)


def _rate_explained(value: float, task: str, vocab: dict) -> bool:
    """率らしい数値が依頼文/用語集のどこかで説明されているか（保守的＝広めに許して
       誤検知を避ける）。①コードの生の数字がそのまま依頼文にある ②%/％表記に換算した
       値が依頼文にある ③用語集の登録値が同じ率を指す、のいずれか。"""
    factor = value if value >= 1 else 1 + value
    frac = value if value < 1 else value - 1
    pct = frac * 100
    if any(lit in task for lit in {f"{value:g}", f"{factor:g}", f"{frac:g}"} if lit):
        return True
    if any(f"{p}%" in task or f"{p}％" in task for p in {f"{pct:g}", f"{pct:.0f}"}):
        return True
    for v in vocab.values():
        try:
            fv = float(v)
        except (TypeError, ValueError):
            continue
        if abs(fv - factor) < 1e-6:
            return True
    return False


def scan_rate_literals(code: str, task: str, vocab: dict | None = None) -> list:
    """生成 Basic コード中の『率らしい数値リテラル』のうち、依頼文にも用語集にも出典が
       無いものを、検算を促す助言のリストにする（同じ数値は1回だけ報告）。"""
    vocab = vocab or {}
    seen: set = set()
    out: list = []
    for line in code.splitlines():
        if _BASIC_COMMENT_LINE_RE.match(line):
            continue
        for m in _RATE_LITERAL_RE.finditer(line):
            raw = m.group(1)
            if raw in seen:
                continue
            value = float(raw)
            if not _looks_like_rate(value):
                continue
            if _rate_explained(value, task, vocab):
                continue
            seen.add(raw)
            out.append(f"★ 率らしい数値 ({raw}) が依頼に無いのに使われています — 検算してください")
    return out


# ---------------------------------------------------------------------------
# ★ W10b 項目1/2: 自由生成の関所とヘルパ総なめ検出
#   operator 第6回ブラインド査定（実機）で確定した致命: 自由生成(FREEFORM/OUT_OF_VOCAB)は
#   事後条件チェッカーが無い唯一の経路なのに、成功時は「⚠ 機械保証なし」のソフト警告だけで
#   無条件に適用されていた。修復ループが行き詰まると、依頼と無関係な操作（ヘルパの全呼び出し）
#   を書いて「何か効くかもしれない」を試す暴走（ヘルパ総なめ）が3回連続で再現した。
#   検証できないものは人に確認を返す（忠実度ゲート/破壊の関所と同じ思想）。
# ---------------------------------------------------------------------------

class _FreeformGateAbort(Exception):
    """★ W10b 項目1: run_freeform_plan_step の関所で人が拒否/非対話で確認できなかった時、
       cmd_run_plan まで一気に抜けるための内部シグナル（_confirm_overwrite_or_gate が
       cmd_run_plan の同じフレームで直接 exit code を return しているのと同じ扱いを、
       別関数をまたいでも実現する）。exit_code は呼び出し元がそのまま return すべき値。"""
    def __init__(self, exit_code: int):
        super().__init__(exit_code)
        self.exit_code = exit_code


def _confirm_freeform_apply(a: argparse.Namespace, sweep_warning: str | None = None,
                             step_prefix: str = "") -> int | None:
    """★ W10b 項目1: 自由生成の関所。DSL 経路（②検証→③確認→④codegen→⑤適用→⑥事後条件）
       と違い、自由生成は「変化したか」しか機械確認できず「正しいか」は検証できない唯一の
       経路（事後条件チェッカーが無い）。生成コードを見せた直後・適用の直前に必ず人に確認を
       返す。--allow-freeform が既に立っていれば承知の上として素通し。
       sweep_warning があれば y/N の直前に強調表示する（★ 項目2: ヘルパ総なめの疑い）。
       戻り値: 続行してよければ None、中断すべきなら呼び出し側がそのまま return すべき
       exit code（対話で拒否=1・非対話で確認できない=8）。
       step_prefix は複合計画の段番号表示用（例: "  2段目: "）。単発 FREEFORM では空文字。"""
    if getattr(a, "allow_freeform", False):
        return None
    if sweep_warning:
        print(f"{step_prefix}{sweep_warning}")
    try:
        ans = input(f"{step_prefix}このコードは機械検証できません。適用しますか？ [y/N]: ").strip().lower()
    except EOFError:
        print(f"{step_prefix}この処理を続けるには、以下のいずれかを指定して再実行してください:")
        print(f"{step_prefix}  --allow-freeform  機械検証できないことを承知の上で適用する")
        return 8
    if ans not in ("y", "yes"):
        print(f"{step_prefix}× 中止した")
        return 1
    return None


_HELPER_SUB_DECL_RE = re.compile(r"^\s*Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.M)
_HELPER_CALL_RE = re.compile(r"\bCall\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(")
# ★ 根拠: bench 再現実験（W10b・demo/expense_w10b.xlsx 相当のダミー表で「氏名の列を全部
#   『退職済み』に書き換えて」を実行）で、修復3回目のヘルパ総なめは異なるヘルパ7種
#   （AutoFitColumns/AlignCenter/FormatThousands/VLookupFromTable/PivotSum/SummaryTable/
#   StyleBold）を Call していた一方、正常な単一操作の自由生成は0〜1種しか呼ばない。
#   間の4を閾値に採用（operator の所見「単一操作でCallが4個以上」とも一致）。
_HELPER_SWEEP_THRESHOLD = 4


def _known_helper_names(helper_files: list) -> set:
    """helper_files（load_helpers が返す .bas ファイル一覧）から Sub 名を集める
       （ヘルパ総なめ検出の対象を『実在するヘルパ』だけに絞るため。ユーザー定義の
       Sub まで数えると通常の複数ヘルパ利用と区別できなくなる）。"""
    names: set = set()
    for f in helper_files:
        try:
            names |= set(_HELPER_SUB_DECL_RE.findall(f.read_text(encoding="utf-8")))
        except OSError:
            continue
    return names


def detect_helper_sweep(code: str, helper_names: set) -> str | None:
    """★ W10b 項目2: 自由生成コードが異なるヘルパを閾値(_HELPER_SWEEP_THRESHOLD)以上
       呼んでいたら、依頼と無関係な操作を全部 Call する『ヘルパ総なめ』の疑いを1行で
       返す（無ければ None）。ブロックはしない（誤検知の可能性は残る）— 関所の y/N の
       直前に強調表示し、人の判断を後押しする助言止まり。"""
    if not helper_names:
        return None
    called = set(_HELPER_CALL_RE.findall(code)) & helper_names
    if len(called) < _HELPER_SWEEP_THRESHOLD:
        return None
    names = "、".join(sorted(called))
    return (f"🚨 疑わしい: {len(called)} 種類のヘルパを呼んでいます（{names}）"
            "— 依頼と無関係な操作が混じっていないか確認してください")


def cmd_run_freeform(a: argparse.Namespace, book: Path, source_book: Path) -> int:
    """自由生成経路（従来の cmd_run 本体そのまま。M2a の助言つき）。
       ① 翻訳が CLARIFY にも DSL 語彙にも決まらなかった（FREEFORM・翻訳失敗）ときに使う。
       ★ W3: source_book は cmd_run が翻訳より前に正規化済み（--dry のときは book と同じ・
       正規化していない）。ここでは正規化をやり直さない。"""
    refs_dir = Path(a.refs).resolve() if a.refs else DEFAULT_REFS
    helpers_dir = Path(a.helpers).resolve() if a.helpers else DEFAULT_HELPERS
    helper_catalog, helper_files = load_helpers(helpers_dir)
    known_helper_names = _known_helper_names(helper_files)   # ★ W10b 項目2: 総なめ検出用
    system = CONTRACT + load_refs(refs_dir) + helper_catalog
    desc = describe_book(book)
    user = f"{desc}\n\nタスク:\n{a.task}\n\n`Sub Run(oDoc As Object)` を1つだけ書け。コードのみ。"
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]

    # ★ W8a 項目5: 「自由生成経路」→「AI が直接作成（機械保証なし）」（operator の
    #   検品リストの語彙翻訳。内部の変数名・関数名・コメントは不変）。
    print(f"■ ailine（AI が直接作成・機械保証なし）  model={a.model}  book={book.name}")
    print(f"■ 参照ライブラリ: {refs_dir}  ({len(list(refs_dir.glob('*.bas'))) if refs_dir.is_dir() else 0} 例)")
    print(f"■ ヘルパ: {helpers_dir}  ({len(helper_files)} 本を同梱・Call で呼ばせる)")

    workdir = book.parent / f".ailine_{book.stem}"
    workdir.mkdir(exist_ok=True)
    out_book = book.with_name(book.stem + ".out" + book.suffix)

    apply_timeout = a.timeout if a.timeout else None   # 0 で無効化（旧挙動 = 無制限）

    before = None if a.dry else snapshot(source_book)
    vocab = load_vocab()   # ★ W8a 項目4: 率リテラルスキャンが「用語集に説明があるか」を見る

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

        # ★ W10b 項目1: 自由生成の関所。コードは既に表示済み（上の attempt 表示）・
        #   適用の直前に必ず確認する（原本にはまだ何も触れていない）。
        sweep_warning = detect_helper_sweep(code, known_helper_names)
        gate_exit = _confirm_freeform_apply(a, sweep_warning)
        if gate_exit is not None:
            return gate_exit

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
                     {"role": "user", "content": f"実行時エラー: {err}\nこれを直して。{_REPAIR_SCOPE_GUARD}コードのみ。"}]
            continue

        after = snapshot(out_book)
        changed, lines = diff_snapshots(before, after)
        if not changed:
            print("× no-op（実行は成功したが文書に変化が無い）。修復する。")
            failure_kind = "noop"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "実行は成功したが文書に一切変化が無かった（no-op）。"
                      "設定した API が効いていない可能性がある。別の正しい方法で書き直して。"
                      f"{_REPAIR_SCOPE_GUARD}コードのみ。"}]
            continue

        # ★ W8a 項目4: 単段の FREEFORM/OUT_OF_VOCAB は、成功しても『機械検証済み』の
        #   ✓ ではなく複数段計画の語彙外段(format_plan_report)と同じ強度の正直な ⚠ 枠で
        #   表示する（実測: 8% 仮定やラベル貼りが「✓できました」で素通りしていた）。
        print("\n⚠ AI が直接作成した処理です（機械保証なし）— 確認してください。変更点:")
        for ln in lines:
            print(ln)
        # ★ 止血3: FREEFORM 経路は no-op ガード/advisories も snapshot() 頼みなので、
        #   検証自体が先頭1000行までしか見ていない（DSL経路より弱い正直さ）。
        notice = _truncation_notice(before, after, exhaustive_postcondition=False)
        if notice:
            print(notice)
        advisories = build_advisories(a.task, before, after)
        # ★ W8a 項目4: 率らしい数値リテラルの機械スキャン（判断棚から昇格）。
        advisories = advisories + scan_rate_literals(code, a.task, vocab)
        for adv in advisories:
            print(adv)
        result["ok"] = True
        result["changes"] = lines
        result["advisories"] = advisories
        failure_kind = "none"
        # ★ W8b-2 項目1: 自由生成(FREEFORM/OUT_OF_VOCAB)は機械保証が無いので、既定(原本
        #   直接適用)でも trailing メッセージは ⚠「機械保証はありません」側を使う。
        _finish_apply(a, book, out_book, workdir, result, machine_verified=False)
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
                            apply_timeout: float | None, step_prefix: str = "") -> tuple:
    """M2c: 複合計画の語彙外(OUT_OF_VOCAB/FREEFORM)段を FREEFORM 経路で実行する。
       cmd_run_freeform と同じ生成→（★ W10b: 関所→）適用→署名/切断/no-op チェックのループを、
       『その段の依頼文だけ』かつ『out_book の現在の状態』を起点に行う版。
       ★ cmd_run_freeform 本体は変えない（既存の回帰リスクを避けるため意図的に複製する）。
       ★ W10b 項目1: 関所で人が拒否/非対話で確認できなかった場合は _FreeformGateAbort を
       投げて cmd_run_plan まで一気に抜ける（破壊の関所と同じ『計画全体を止める』扱い。
       原本(book)はこの時点でまだ一切触れていない＝out_book はコピーなので安全）。
       step_prefix は複合計画の段番号表示用（例: "  2段目: "）。
       戻り値: (ok, changes:list[str], advisories:list[str], failure_kind:str, detail:str|None)"""
    helper_catalog, helper_files = load_helpers(helpers_dir)
    known_helper_names = _known_helper_names(helper_files)   # ★ W10b 項目2: 総なめ検出用
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

        # ★ W10b 項目1: 自由生成の関所。単発 cmd_run_freeform と違い、この経路はこれまで
        #   生成コードを一切表示していなかった（黙って確認を求めても判断できない）ので、
        #   ここで初めて表示してから y/N を聞く。
        print(f"{step_prefix}─ 生成した .bas（語彙外・AI が直接作成）───────────────")
        print(code)
        print(f"{step_prefix}──────────────────────────────────────────")
        sweep_warning = detect_helper_sweep(code, known_helper_names)
        gate_exit = _confirm_freeform_apply(a, sweep_warning, step_prefix=step_prefix)
        if gate_exit is not None:
            raise _FreeformGateAbort(gate_exit)

        shutil.copy2(stepsource, out_book)
        t0 = progress_start("⏳ LibreOffice で適用中…")
        ok, err, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not ok:
            failure_kind = "runtime_error"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": f"実行時エラー: {err}\nこれを直して。{_REPAIR_SCOPE_GUARD}コードのみ。"}]
            continue

        after = snapshot(out_book)
        changed, lines = diff_snapshots(before, after)
        if not changed:
            failure_kind = "noop"
            msgs += [{"role": "assistant", "content": raw},
                     {"role": "user", "content":
                      "実行は成功したが文書に一切変化が無かった（no-op）。"
                      "設定した API が効いていない可能性がある。別の正しい方法で書き直して。"
                      f"{_REPAIR_SCOPE_GUARD}コードのみ。"}]
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
                # ★ W8a 項目5: 「自由生成」→「AI が直接作成（機械保証なし）」（operator の語彙翻訳）。
                lines.append(f"{idx}. {label} → {mark} 語彙外のため AI が直接作成（機械保証なし）で実行（確認してください）")
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
        # ★ W8a 項目5: 表示文言のみ「自由生成」→「AI が直接作成」（operator の語彙翻訳）。
        return ("⚠ 一部は確認が必要です（語彙外で AI が直接作成した段、または検証対象不足の段があり、"
                "機械検証はしていません）", "warn")
    return "✓ すべて機械検証済み", "ok"


# ★ W10d【本命】: 複合計画は段が増えるほど同じ助言（例: 幽霊データ検出）が段の数だけ
#   並びうる。査定で名指しされたオオカミ少年化を避けるため、文言が同じものは1行に畳んで
#   段番号を列挙する（初出順は保つ・文言そのものは変えない＝読み手が信じる根拠を削らない）。
def _group_step_advisories(entries: list) -> list:
    """entries: [(段番号 or None, 助言文言), ...]（None=計画全体に対する助言・段に紐付かない）
       を文言でグルーピングし、初出順を保った [(段番号のリスト, 文言), ...] を返す。
       印字用(_dedup_step_advisories)と --json 用(cmd_run_plan の result["advisories"])が
       同じグルーピングを共有する（二重管理しない）。"""
    order = []
    by_text: dict = {}
    for idx, text in entries:
        if text not in by_text:
            by_text[text] = []
            order.append(text)
        if idx is not None and idx not in by_text[text]:
            by_text[text].append(idx)
    return [(by_text[text], text) for text in order]


def _dedup_step_advisories(entries: list) -> list:
    """複合計画の全段から集めた助言を、印字用の行リストに整形する（同じ文言は1行に畳み、
       該当する段番号を列挙する）。"""
    lines = []
    for idxs, text in _group_step_advisories(entries):
        if idxs:
            prefix = "・".join(f"{i}段目" for i in idxs)
            lines.append(f"  {prefix}: {text}")
        else:
            lines.append(f"  {text}")
    return lines


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
        print("\n（--dry プレビュー・語彙外の段は実行時に AI が直接作成（機械保証なし）で対応します。未実行）")
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
    step_advisory_entries: list = []   # ★ W10d: [(段番号 or None, 助言文言), ...]
    mention_exclude_sheets: set = set()   # ★ W10d: LOOKUP_FILL の参照専用シート（全段分の合算）

    for i, step in enumerate(plan, 1):
        op = step.get("op")

        if op == "CLARIFY":
            question = step.get("question") or "確認が必要です"
            items.append((i, question, "fail", "計画の途中で確認が必要なため対応できません"))
            plan_json.append({"op": "CLARIFY", "command": None, "status": "fail", "postcondition": None})
            continue

        if op not in OP_SCHEMA:
            about = step.get("about") or "内容不明の依頼"
            # ★ W10b 項目1: 関所で拒否/非対話なら、破壊の関所と同じく計画全体をここで止める
            #   （原本(book)はまだ無傷・out_book はコピーなので途中終了しても安全）。
            try:
                okf, changes, advisories, _fkind, detail = run_freeform_plan_step(
                    a, about, out_book, workdir, refs_dir, helpers_dir, f"plan{i}", apply_timeout,
                    step_prefix=f"  {i}段目: ")
            except _FreeformGateAbort as e:
                return e.exit_code
            if okf:
                items.append((i, about, "warn", None))
                for ln in changes:
                    print(f"  {ln}")
                # ★ W10d: 段ごとに即印字せず、他段の助言と合わせてループの後で
                #   重複を畳んでから出す（run_freeform_plan_step は既にこの段の
                #   依頼文(about)だけを見た build_advisories を返しており、局所判定のまま流用できる）。
                step_advisory_entries.extend((i, adv) for adv in advisories)
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
        if op == "PIVOT":   # ★ W9 項目4
            print(f"  {i}段目: （{PIVOT_CAVEAT}）")
        warn_overwrite = _maybe_warn_target_overwrite(op, resolved, current_meta, out_book)
        if warn_overwrite:
            summary = _interpretation_summary_line(resolved, inferred)   # ★ W10a 項目3
            if summary:
                print(f"  {i}段目: {summary}")
            print(f"  {i}段目: {warn_overwrite}")
        for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
            print(f"  {i}段目: ⚠ {w}")
        # ★ W10a 項目1: 破壊の関所（複合計画の段ごと）。原本にはまだ何も反映されていない
        #   （out_book はコピー・最終段まで揃ってから atomic_replace_inplace される）ので、
        #   ここで中断すれば原本は無傷のまま run 全体を止められる。
        gate_exit = _confirm_overwrite_or_gate(a, warn_overwrite, step_prefix=f"  {i}段目: ")
        if gate_exit is not None:
            return gate_exit
        if resolved.get("_sources"):
            plan_provenance.append({"step": i, **resolved["_sources"]})
        step_header_row = current_meta.get("header_rows", {}).get(first_sheet, 1) if first_sheet else 1
        code = codegen_dsl(op, resolved, current_meta, use_formula=use_formula)
        (workdir / f"plan_step{i}.bas").write_text(code, encoding="utf-8")

        # ★ W9: INSERT_ROWS/AUTOFIT の事後条件が段ごとの before/after を突き合わせられる
        #   よう、この段の適用直前の out_book をコピーして残す（run_freeform_plan_step と
        #   同じ考え方・他 op には無害な余分なコピー1回）。
        stepsource = workdir / f"plan_step{i}_source{out_book.suffix}"
        shutil.copy2(out_book, stepsource)
        step_before = snapshot(stepsource)   # ★ W10d: 助言計算用（この段の適用直前）

        t0 = progress_start(f"⏳ {i}段目 LibreOffice で適用中…")
        okrun, err_apply, _raw = basrun_apply(out_book, code, workdir, helper_files, timeout=apply_timeout)
        progress_end(t0)
        if not okrun:
            detail = f"実行時エラー: {short_error_summary(err_apply)}"
            items.append((i, label, "fail", detail))
            plan_json.append({"op": op, "command": line, "status": "fail", "postcondition": None})
            continue

        # ★ W10d【本命】: 単発 op(cmd_run_dsl)なら出る助言（幽霊データ/一様埋め/件数突き合わせ/
        #   新規シート中身・申告）が、複合計画の DSL 段では丸ごと欠落していた（build_advisories
        #   を一度も呼んでいなかった＝前任の W10c 報告の未処置の穴そのもの）。
        #   ★ 依頼文言との重なり(mention_overlap_advisory)はここに含めない: 段ごとの
        #   before/after だけでは複合依頼全体に対する充足を判定できない（他段が担当する
        #   言及まで『この段で変更されていない』と誤検知する＝ W6 でシート混在に見つかった
        #   ものと同じ「全部該当」崩れの再演）。ループの外で before_all/after_all に対して
        #   一度だけ評価する。
        step_after = snapshot(out_book)
        if op == "LOOKUP_FILL" and resolved.get("source_sheet"):
            mention_exclude_sheets.add(resolved["source_sheet"])
        step_adv = _structural_advisories(step_before, step_after)
        step_adv.extend(unrequested_new_sheet_advisory(a.task, step_before, step_after))
        step_adv = _neutralize_new_column_ghost_warning(step_adv, op, resolved, current_meta)
        step_adv = _neutralize_declared_new_sheet_warning(step_adv, op, step_before, step_after)
        step_advisory_entries.extend((i, adv) for adv in step_adv)

        status, reason = run_postcondition(op, out_book, resolved, before_charts=before_charts,
                                            header_row=step_header_row, use_formula=use_formula,
                                            source_book=stepsource)
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

    # ★ W10d【本命】: 依頼文言との重なり(④ mention_overlap_advisory)は計画全体に対して
    #   一度だけ評価する（段ごとの局所判定では『他段が担当する言及』を誤検知するため・
    #   _structural_advisories のコメント参照）。exclude_sheets は全段の LOOKUP_FILL
    #   source_sheet を合算する。
    after_all = snapshot(out_book)
    mentions = extract_task_mentions(a.task, before_all["sheets"])
    final_mention_lines = mention_overlap_advisory(
        mentions, before_all, after_all, mention_exclude_sheets or None)
    step_advisory_entries.extend((None, ln) for ln in final_mention_lines)   # None=計画全体

    print()
    for ln in format_plan_report(items):
        print(ln)
    dedup_advisories = _dedup_step_advisories(step_advisory_entries)   # ★ W10d: 重複を畳む
    if dedup_advisories:
        print("\n助言:")
        for ln in dedup_advisories:
            print(ln)
    verdict_line, verdict = overall_verdict(items)
    print(f"\n{verdict_line}")

    _changed, difflines = diff_snapshots(before_all, after_all)
    result["plan"] = plan_json
    result["provenance"] = plan_provenance or None
    result["items"] = [{"idx": idx, "label": label, "status": st, "detail": det}
                        for idx, label, st, det in items]
    result["changes"] = difflines
    result["advisories"] = [{"steps": idxs, "text": text}
                             for idxs, text in _group_step_advisories(step_advisory_entries)]

    if verdict == "fail":
        result["out"] = str(out_book)
        _finish_run(a, book, result, "plan_step_failed")
        return 1

    result["ok"] = True
    # ★ W8b-2 項目1: 複合計画は総合判定(overall_verdict)に従う。全段機械検証済み(ok)
    #   の時だけ ✓「反映しました」、語彙外/検証不足の段が混じる(warn)なら
    #   ⚠「機械保証はありません」側（自由生成の段が混じっている以上、全体としても
    #   機械保証済みとは名乗れない＝format_plan_report/overall_verdict と同じ誠実さ）。
    _finish_apply(a, book, out_book, workdir, result, machine_verified=(verdict == "ok"))

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
    r.add_argument("--inplace", action="store_true",
                   help="（廃止・後方互換のため受理のみ）既定で原本に直接適用するため不要。"
                        "旧 .out 挙動が欲しければ --copy")
    r.add_argument("--json", action="store_true", help="結果を JSON でも出す")
    r.add_argument("--timeout", type=float, default=DEFAULT_APPLY_TIMEOUT,
                   help=f"basrun apply のタイムアウト秒 (既定 {DEFAULT_APPLY_TIMEOUT:.0f}、"
                        "0 で無効化=旧挙動の無制限)")
    r.add_argument("--ask", action="store_true",
                   help="DSL 経路の確認行の後に y/n で対話する（既定は表示して続行）")
    r.add_argument("--keep-backups", dest="keep_backups", type=int, default=DEFAULT_KEEP_BACKUPS,
                   help=f"原本への反映前のバックアップを book ごとに何世代残すか (既定 {DEFAULT_KEEP_BACKUPS}、"
                        "負数で無制限)")
    r.add_argument("--values", action="store_true",
                   help="COMPUTE_COLUMN を式でなく値ベタ書きにする（既定は式・W3 Part3）")
    r.add_argument("--header-row", dest="header_row", type=int, default=None,
                   help="見出し行を明示指定（1起点。指定時は自動検出をスキップしてこの行を採用）")
    r.add_argument("--accept-loss", dest="accept_loss", action="store_true",
                   help="往復忠実度ゲートが検出した喪失を承知の上で原本への反映を続行する"
                        "（ailine undo で戻せる）")
    r.add_argument("--copy", action="store_true",
                   help="原本には触らず <book>.out に結果を作る（既定の原本直接適用をしない・"
                        "旧 --inplace 無指定と同じ挙動。往復忠実度ゲートも走らせない）")
    r.add_argument("--overwrite", action="store_true",
                   help="破壊の関所（既存データを持つ列への上書き確認）を承知の上で"
                        "続行する（ailine undo で戻せる）")
    r.add_argument("--allow-freeform", dest="allow_freeform", action="store_true",
                   help="自由生成の関所（AI が直接作成したコードは機械検証できないという確認）を"
                        "承知の上で続行する（ailine undo で戻せる）")
    r.set_defaults(func=cmd_run)

    s = sub.add_parser("stop", help="起動した LibreOffice を落とす")
    s.set_defaults(func=cmd_stop)

    d = sub.add_parser("doctor", help="セットアップを診断する")
    d.add_argument("--model", default=DEFAULT_MODEL, help=f"確認するモデル (既定 {DEFAULT_MODEL})")
    d.set_defaults(func=cmd_doctor)

    h = sub.add_parser("history", help="実行履歴を表示する")
    h.add_argument("--max", type=int, default=10, help="表示件数（既定 10、新しい順）")
    h.set_defaults(func=cmd_history)

    rs = sub.add_parser("restore", help="原本への反映前のバックアップから復元する（ailine undo と同じ）")
    rs.add_argument("book", help="対象の文書 (.xlsx / .ods)")
    rs.add_argument("--list", action="store_true", help="バックアップ一覧を表示するだけ（復元しない）")
    rs.set_defaults(func=cmd_restore)

    u = sub.add_parser("undo", help="原本への反映前のバックアップから復元する（あと何回戻せるかを表示）")
    u.add_argument("book", help="対象の文書 (.xlsx / .ods)")
    u.add_argument("--list", action="store_true", help="バックアップ一覧を表示するだけ（復元しない）")
    u.set_defaults(func=cmd_undo)

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
