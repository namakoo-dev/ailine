"""基本操作 × 言い回しの揺れ × 表の形 ── **効果**で測る検体。

★★ なぜ要るか（2026-08-29・Namakoo「2 日前から同様の事例を繰り返している」）:
  この 2 日、壊れは全部**人が手で触って**見つかった。俺は分母を持たずに 1 件ずつ
  直していた ── 直った数は分かるが、**残りがどれだけあるか**を誰も知らない。
  ★ 既に在る凍結検体（translation_battery）は **op 名の一致**しか見ない。
    op が合っていても「担当列が全行書き換わる」事故は起きるし、op が違っても
    結果が正しいことはある（読み直しの層が拾う）。**測るべきは効果**。

測り方: 表を作る → 依頼文を投げる → **出来上がったファイルを読んで**、
        意図した効果になっているかだけを見る（op 名も内部の経路も見ない）。

  合格 = 期待した変化がちょうど起きている（かつ他が壊れていない）
  断り = ？ で止まった（**壊していない**ので、間違いより軽い ── 別枠で数える）
  失敗 = 何か違うことをした（一番重い）

使い方:
    python bench/basic_ops_matrix.py                 # 全部
    python bench/basic_ops_matrix.py --table 在庫    # 表を絞る
    python bench/basic_ops_matrix.py --op cell       # 操作を絞る
    python bench/basic_ops_matrix.py --out result.json

★ 実物の ollama + LibreOffice が要る（遅い）。CI では走らせない。
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "src"))

# --- 表（語彙も列名も形も違う 3 つ）-----------------------------------------------------

TABLES = {
    "在庫": {
        "sheet": "在庫",
        "headers": ["品名", "棚", "数量", "備考"],
        "rows": [["ボルト", "A-1", 120, ""], ["ナット", "A-2", 80, ""],
                  ["ワッシャー", "B-1", 300, ""]],
    },
    "名簿": {
        "sheet": "名簿",
        "headers": ["氏名", "所属", "内線", "メモ"],   # ★ 4 列目はわざと見慣れない語
        # ★ 2026-08-29 の 1 回目で分かった検体の欠陥: 内線を 101/202/303 にすると
        #   303 = 101+202 になり、**合計行の番人が最下行を合計と誤検出**して並べ替えが
        #   × になった（番人は正しく疑っている ── 悪いのは検体）。
        #   ★ 検体の数字が偶然の関係を持たないようにする。
        "rows": [["山田", "営業", 101, ""], ["鈴木", "経理", 202, ""],
                  ["高橋", "総務", 305, ""]],
    },
    "献立": {
        "sheet": "献立",
        "headers": ["料理", "主材料", "分量", "備考"],
        "rows": [["カレー", "牛肉", 4, ""], ["味噌汁", "豆腐", 2, ""],
                  ["サラダ", "レタス", 3, ""]],
    },
}


def _build(table_key: str, path: Path) -> Path:
    t = TABLES[table_key]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = t["sheet"]
    ws.append(t["headers"])
    for r in t["rows"]:
        ws.append(list(r))
    wb.save(path)
    return path


# --- 期待する効果（op 名でなく、出来上がったファイルで見る）----------------------------

def _grid(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    return [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
             for r in range(1, ws.max_row + 1)]


def cell_becomes(row: int, col: int, value):
    """その 1 セルだけが value になり、**他は 1 セルも変わっていない**こと。"""
    def check(before, after):
        if len(before) != len(after):
            return False, f"行数が変わった {len(before)}→{len(after)}"
        diff = [(r, c) for r in range(len(after)) for c in range(len(after[r]))
                 if before[r][c] != after[r][c]]
        if diff != [(row - 1, col - 1)]:
            return False, f"変わったセルが {diff}（{row}行{col}列だけのはず）"
        got = after[row - 1][col - 1]
        if str(got) != str(value):
            return False, f"値が {got!r}（{value!r} のはず）"
        return True, ""
    return check


def row_added_at(at: int, first_value):
    """at 行目に 1 行増え、1 列目がその値。**他の行は順序ごとそのまま**。"""
    def check(before, after):
        if len(after) != len(before) + 1:
            return False, f"行数が {len(before)}→{len(after)}（1 増えるはず）"
        got = after[at - 1][0]
        if str(got) != str(first_value):
            return False, f"{at}行目の 1 列目が {got!r}（{first_value!r} のはず）"
        rest = after[:at - 1] + after[at:]
        if rest != before:
            return False, "他の行が変わっている"
        return True, ""
    return check


def row_deleted(name):
    """その名前の行が消え、**他は順序ごとそのまま**。"""
    def check(before, after):
        want = [r for r in before if str(r[0]) != str(name)]
        if len(after) != len(before) - 1:
            return False, f"行数が {len(before)}→{len(after)}（1 減るはず）"
        if after != want:
            return False, "消えた行が違う（または他の行が動いた）"
        return True, ""
    return check


def column_added_named(name, at_index=None):
    """見出しに name の列が増える（位置指定があればその位置）。値は増やさない。"""
    def check(before, after):
        if len(after) != len(before):
            return False, f"行数が変わった {len(before)}→{len(after)}"
        heads = [str(h or "") for h in after[0]]
        if name not in heads:
            return False, f"見出しに『{name}』が無い（{heads}）"
        if at_index is not None and heads.index(name) + 1 != at_index:
            return False, f"『{name}』が {heads.index(name) + 1} 列目（{at_index} 列目のはず）"
        return True, ""
    return check


def column_deleted(name):
    def check(before, after):
        heads_b = [str(h or "") for h in before[0]]
        heads_a = [str(h or "") for h in after[0]]
        if name in heads_a:
            return False, f"『{name}』がまだ在る"
        if len(heads_a) != len(heads_b) - 1:
            return False, f"列数が {len(heads_b)}→{len(heads_a)}（1 減るはず）"
        return True, ""
    return check


def rows_sorted_by(col_index: int, desc: bool):
    def check(before, after):
        if len(after) != len(before):
            return False, "行数が変わった"
        vals = [r[col_index - 1] for r in after[1:]]
        want = sorted(vals, reverse=desc)
        if vals != want:
            return False, f"並びが {vals}（{want} のはず）"
        return True, ""
    return check


# --- 検体（基本操作 × 言い回し）---------------------------------------------------------
#
# ★ 言い回しは「思いつく限り」ではなく、**人が実際に書きそうな形**を並べる。
#   ここが痩せていると「高精度」の主張が空になる。

def _cases_for(key: str):
    t = TABLES[key]
    h = t["headers"]
    r1, r2, r3 = (row[0] for row in t["rows"])
    c1 = h[0]           # 1 列目の見出し
    c2 = h[1]           # 2 列目
    c3 = h[2]           # 3 列目（数値）
    out = []

    # ① セルに値を入れる（1 セルだけ）
    for task in (f"{r2}の{c2}を「東棟」にして",
                  f"{r2}の{c2}に東棟と入れて",
                  f"{r2}の{c2}を東棟に変えて",
                  f"3行目の{c2}を「東棟」にして",
                  f"{r2}の右に東棟",
                  f"{r2}の隣に東棟"):
        out.append(("cell", task, cell_becomes(3, 2, "東棟")))
    out.append(("cell", f"{r3}の{c3}を「999」にして", cell_becomes(4, 3, 999)))
    out.append(("cell", f"4行目の{c3}を999にして", cell_becomes(4, 3, 999)))

    # ② 行を足す（名前つき・位置指定）
    for task in (f"{r1}と{r2}の間に新品を作って",
                  f"{r1}と{r2}の間に新品を追加して",
                  f"{r1}と{r2}の間に新品を入れて",
                  f"{r1}の下に新品を追加して"):
        out.append(("row_add", task, row_added_at(3, "新品")))
    out.append(("row_add", f"{r3}の下に新品を追加して", row_added_at(5, "新品")))
    out.append(("row_add", f"{r2}の上に新品を入れて", row_added_at(3, "新品")))
    # ★★ 2026-08-29（Namakoo）:「どうしても中身でさせない場面が出てくる。例えば
    #   4行目と5行目は両方ともヤマノ食品。取引先で指定は出来ない」
    #   ★ 同じ値が 2 行あれば中身では指せない ── **番号でしか言えない場面がある**。
    #   実測で、番号で言うと 4 行目に空行が挿さっていた（下ではなく上・値も入らない）。
    for task, at in ((f"3行目の下に新品を追加して", 4), (f"3行目の上に新品を入れて", 3),
                      (f"3行目と4行目の間に新品を作って", 4)):
        out.append(("row_add", task, row_added_at(at, "新品")))

    # ③ 行を消す
    for task in (f"{r2}の行を削除して", f"{r2}を削除して", f"{r2}の行を消して",
                  f"{r2}の行を除いて", "3行目を削除して"):
        out.append(("row_del", task, row_deleted(r2)))

    # ④ 列を足す
    for task in (f"{c2}の右に区分という列を追加して",
                  f"{c2}の右に区分の列を作って",
                  f"{c1}と{c2}の間に区分の列を追加して"):
        idx = 3 if "右" in task else 2
        out.append(("col_add", task, column_added_named("区分", idx)))
    out.append(("col_add", "区分という列を追加して", column_added_named("区分")))

    # ⑤ 列を消す
    for task in (f"{c2}の列を削除して", f"{c2}列を消して"):
        out.append(("col_del", task, column_deleted(c2)))

    # ⑥ 並べ替え
    out.append(("sort", f"{c3}で降順に並べ替えて", rows_sorted_by(3, True)))
    out.append(("sort", f"{c3}の大きい順にして", rows_sorted_by(3, True)))
    out.append(("sort", f"{c3}の少ない順に並べて", rows_sorted_by(3, False)))
    return out


def run_one(table_key: str, task: str, check, workdir: Path, timeout: float):
    src = _build(table_key, workdir / "in.xlsx")
    sheet = TABLES[table_key]["sheet"]
    before = _grid(src, sheet)
    out = workdir / "in.out.xlsx"
    if out.exists():
        out.unlink()
    env = {**__import__("os").environ, "PYTHONPATH": str(ROOT / "src")}
    # ★ 2026-08-30: 入口を差し替えられるようにした（AILINE_ENTRY）。
    #   既定は製品そのもの。段階的に聞く実験（bench/staged_translate.py）を
    #   **同じ検体で**測るためだけの口で、製品の挙動は 1 ビットも変わらない。
    _entry = __import__("os").environ.get("AILINE_ENTRY")
    _cmd = ([sys.executable, _entry] if _entry else [sys.executable, "-m", "ailine"])
    p = subprocess.run(
        _cmd + ["run", str(src), task, "--copy",
                 "--sheet", sheet, "--timeout", "90"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout, cwd=str(ROOT), env=env)
    stdout = p.stdout or ""
    if p.returncode == 3 or "？" in stdout.split("\n")[-6:] and p.returncode != 0:
        return "refused", _first_line(stdout, "？")
    if p.returncode != 0:
        return "failed", _first_line(stdout, "×") or f"exit {p.returncode}"
    if not out.exists():
        return "failed", "出力が無い"
    ok, why = check(before, _grid(out, sheet))
    return ("pass", "") if ok else ("failed", why)


def _first_line(text: str, mark: str) -> str:
    for ln in (text or "").split("\n"):
        s = ln.strip()
        if s.startswith(mark):
            return s[:160]
    return ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--table", default=None, choices=sorted(TABLES))
    ap.add_argument("--op", default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--timeout", type=float, default=420)
    ns = ap.parse_args()

    keys = [ns.table] if ns.table else sorted(TABLES)
    results, tally = [], {"pass": 0, "refused": 0, "failed": 0}
    with tempfile.TemporaryDirectory(prefix="ailine_matrix_") as tmp:
        work = Path(tmp)
        for key in keys:
            for op, task, check in _cases_for(key):
                if ns.op and op != ns.op:
                    continue
                try:
                    verdict, why = run_one(key, task, check, work, ns.timeout)
                except subprocess.TimeoutExpired:
                    verdict, why = "failed", "時間切れ"
                tally[verdict] += 1
                results.append({"table": key, "op": op, "task": task,
                                 "verdict": verdict, "why": why})
                mark = {"pass": "✓", "refused": "？", "failed": "×"}[verdict]
                print(f"{mark} [{key}/{op}] {task}" + (f"  ── {why}" if why else ""))

    n = sum(tally.values())
    print()
    print(f"合計 {n} 件: ✓ {tally['pass']}  ？断り {tally['refused']}  × 失敗 {tally['failed']}")
    if n:
        print(f"  意図どおり {tally['pass'] / n * 100:.1f}%  "
               f"／ 壊していない（✓+断り） {(tally['pass'] + tally['refused']) / n * 100:.1f}%")
    if ns.out:
        Path(ns.out).write_text(json.dumps(
            {"tally": tally, "results": results}, ensure_ascii=False, indent=2),
            encoding="utf-8")
        print(f"  → {ns.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
