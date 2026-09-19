# -*- coding: utf-8 -*-
"""断りが示した道を**実際に歩く**（2026-09-19・導通試験）。

★★ なぜ要るか（Namakoo「到達率を極大にするのは大事だけど、**漏らした依頼を適切に判断し
  ユーザを正解に導いてやること**までやるから」）: 合格線の 3 条目「通る道を示す」は、
  機械が **escape（逃げ道の語が本文に在る）/ example（例が在る）** しか見ていなかった。
  **その道が通るか**は誰も確かめていない ── 今日ずっと潰してきた「宣言 vs 実体」の隙間。

★★ 判定は 5 つ（Namakoo 承認・2026-09-19）:

    walked      示した道を歩いたら着いた
    by_design   意図して行き止まり（パスの打ち間違い等・unlock が「無い」）
    vague       ★ 理由が粗くて道を示せていない（直せる見込みが高い）
    no_path     本当に道が無い（機能・語彙が無い）
    path_fails  ★★ 歩いたが着かなかった ── **通らない道を示した**（最悪）

  ★ `path_fails` は `no_path` より重い。道が無いのは不足だが、通らない道を示すのは誤情報。
  ★ 道具を作る前に 1 件を手で歩いた時点で `path_fails` が 1 件出ている
    （`--op ADD_ROW` → 「その列のデータ行を全部書き換えます」＝ ADD_ROW には嘘）。

★ verdict は台帳に**書かない** ── 歩いた結果から決める（名簿を写して名簿と比べない）。
★ ここは LLM を回さない群（A 群）だけを歩く。翻訳を固定して引き金を引くので
  **素の環境でも毎回走る**（揺れが混ざらないので判定が安定する）。
  B 群（実機が要る 10 件）は -m local 側の仕事。

★ なぜ tests/ に在るか: 素の環境の番人は scripts/ 同士の import を弾く（3 度踏んだ）。
"""
from __future__ import annotations

import argparse
import contextlib
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402

REGISTER = REPO / "tests" / "refusal_register.json"

#: 歩く時に使う冊（フォルダ経路なので 2 枚置く）。★ 固定する ── 冊が変わると引き金が動く。
BOOKS = {"4月.xlsx": [["商品", "金額"], ["ボルト", 120], ["ナット", 80]],
         "5月.xlsx": [["商品", "金額"], ["ボルト", 150], ["ワッシャー", 300]]}

#: 同じ見出しのシートを 2 枚持つ冊（対象シートが決まらない断りの引き金）。
TWO_SHEETS = {"4月": [["商品", "金額"], ["ボルト", 120], ["ナット", 80]],
              "5月": [["商品", "金額"], ["ボルト", 150], ["ワッシャー", 300]]}


def _make_book(root: Path, sheets: dict | None = None) -> Path:
    """1 冊の冊を作る（sheets を渡せば複数シート）。"""
    root.mkdir(parents=True, exist_ok=True)
    p = root / "在庫.xlsx"
    wb = openpyxl.Workbook()
    if sheets:
        wb.remove(wb.active)
        for name, rows in sheets.items():
            ws = wb.create_sheet(name)
            for r in rows:
                ws.append(r)
    else:
        ws = wb.active
        ws.title = "在庫"
        for r in BOOKS["4月.xlsx"]:
            ws.append(r)
    wb.save(p)
    wb.close()
    return p


def load_register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def _make_folder(root: Path) -> Path:
    d = root / "月次"
    d.mkdir(parents=True, exist_ok=True)
    for name, rows in BOOKS.items():
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in rows:
            ws.append(r)
        wb.save(d / name)
        wb.close()
    return d


def _run(argv: list, plan, second=None) -> tuple:
    """製品を 1 回走らせ、(exit, 画面) を返す。plan を渡せば翻訳をそれに固定する。

    ★ second を渡すと **2 回目の読みだけ別の計画**にする（読みの割れを再現するため）。
    """
    buf = io.StringIO()
    real, real_fixed = ailine.translate_task, ailine.translate_task_fixed_op
    calls = []
    if plan is not None:
        def fake(*a, **k):
            calls.append(1)
            return {"plan": second if (second and len(calls) == 2) else plan}
        ailine.translate_task = fake
        ailine.translate_task_fixed_op = lambda model, op, task, meta, **k: {
            "op": op, "args": dict((plan[0] or {}).get("args") or {})}
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                rc = ailine.main(argv)
            except SystemExit as e:
                rc = e.code
    finally:
        ailine.translate_task, ailine.translate_task_fixed_op = real, real_fixed
    return rc, buf.getvalue()


def _example_in(text: str) -> str | None:
    """断り文が見せている「通る書き方」を取り出す（『…』か「…」の中で、依頼文らしいもの）。"""
    import re
    for m in re.finditer(r"[『「]([^』」]{6,60})[』」]", text):
        s = m.group(1)
        if any(w in s for w in ("抜き出", "並べ替え", "追加", "にして", "消して", "入れて")):
            return s
    return None


def _trigger(w: dict, root: Path) -> tuple:
    """引き金を引く ── (exit, 画面, 引数の組み立てに使う土台)。"""
    kind = w.get("kind")
    extra = list(w.get("argv_extra") or [])
    if kind == "folder":
        folder = _make_folder(root)
        argv = ["run", str(folder), w["task"], "--out", str(root / "結果.xlsx")]
        return (*_run(argv, w.get("plan"), w.get("second")), folder)
    if kind == "two_sheets":
        book = _make_book(root, TWO_SHEETS)
        return (*_run(["run", str(book), w["task"], "--copy"] + extra,
                      w.get("plan"), w.get("second")), book)
    book = _make_book(root)
    return (*_run(["run", str(book), w["task"], "--copy"] + extra,
                  w.get("plan"), w.get("second")), book)


def _walk_path(w: dict, path: dict, base, root: Path) -> tuple:
    """示された道を歩く ── (exit, 画面)。歩けない種類なら (None, 理由)。"""
    kind, extra = path.get("kind"), list(path.get("argv_extra") or [])
    plan = path.get("plan") or w.get("plan")
    if kind == "single_book":
        one = base / next(iter(BOOKS))
        return _run(["run", str(one), w["task"], "--copy"], plan)
    if kind == "example":
        task2 = path.get("task") or _example_in(_LAST_SCREEN[0])
        if not task2:
            return None, "例を示していると数えられているが、歩ける文が取り出せない"
        if w.get("kind") == "folder":
            return _run(["run", str(base), task2, "--out", str(root / "結果2.xlsx")], plan)
        return _run(["run", str(base), task2, "--copy"] + extra, plan)
    if kind == "forced_op":
        return _run(["run", str(base), w["task"], "--copy", "--op", path["op"]], plan)
    if kind == "sheet":
        return _run(["run", str(base), w["task"], "--copy", "--sheet", path["sheet"]], plan)
    return None, f"歩き方が不明: {kind}"


#: ★ 例を断り文から拾う時に使う（直前の画面）。引数で回すと walk_one の形が崩れるので棚に置く。
_LAST_SCREEN = [""]


def walk_one(key: str, entry: dict, root: Path) -> dict:
    """1 件の断りを出させ、示された道を歩く。戻り値に verdict を入れる（台帳は読まない）。"""
    w = entry.get("walk")
    if not w:
        return {"key": key, "verdict": "未記入", "detail": "walk 欄が無い"}
    if w.get("kind") == "match_unwalkable":
        # ★ 突き合わせは 2 冊と実表が要る ── この器では歩けないと**書いて残す**。
        #   「歩いていない」を walked と混ぜない（見ていないものを見たことにしない）。
        return {"key": key, "verdict": "未調査", "detail": w.get("why") or "この器では歩けない"}

    rc, out, base = _trigger(w, root)
    _LAST_SCREEN[0] = out
    if rc == 0:
        return {"key": key, "verdict": "引き金が引けない",
                "detail": "断りが出ずに到達した（検体が古い）", "screen": out[-200:]}

    path = w.get("path") or {}
    if path.get("kind") == "none":
        # ★ 道を示していない ── 理由が粗いのか、本当に道が無いのかは人が仕分ける。
        return {"key": key, "verdict": "vague" if w.get("why") else "no_path",
                "detail": w.get("why") or "示す道が無い（機能・語彙が無い）",
                "screen": out.strip()[-160:]}

    rc2, out2 = _walk_path(w, path, base, root)
    if rc2 is None:
        return {"key": key, "verdict": "vague", "detail": out2, "screen": out.strip()[-160:]}
    if rc2 == 0:
        return {"key": key, "verdict": "walked",
                "detail": f"道を歩いて到達（{path.get('kind')}）"}
    return {"key": key, "verdict": "path_fails",
            "detail": f"道を歩いたが exit {rc2}", "screen": out2.strip()[-200:]}


def survey() -> list:
    reg = load_register()
    rows = []
    with tempfile.TemporaryDirectory() as td:
        for key, entry in reg["refusals"].items():
            if str(entry.get("unlock", "")).startswith("無い"):
                rows.append({"key": key, "verdict": "by_design",
                             "detail": str(entry["unlock"])[:60]})
                continue
            if not entry.get("walk"):
                rows.append({"key": key, "verdict": "未記入", "detail": "walk 欄が無い"})
                continue
            rows.append(walk_one(key, entry, Path(td) / key.replace("#", "_")))
    return rows


ORDER = ["path_fails", "vague", "no_path", "未調査", "未記入", "引き金が引けない",
         "walked", "by_design"]


def render(rows: list) -> str:
    from collections import Counter
    c = Counter(r["verdict"] for r in rows)
    lines = ["導通の盤 ── 断りが示した道を歩く", ""]
    for v in ORDER:
        if c.get(v):
            lines.append(f"  {v:<16} {c[v]:>3}")
    lines.append(f"  {'合計':<16} {len(rows):>3}")
    lines.append("")
    for v in ORDER:
        hits = [r for r in rows if r["verdict"] == v and v != "by_design"]
        if not hits:
            continue
        lines.append(f"■ {v}")
        for r in hits:
            lines.append(f"    {r['key']:<28} {r['detail']}")
            if r.get("screen"):
                lines.append(f"        画面: {r['screen'][:110]}")
        lines.append("")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="断りが示した道を歩く（導通試験）")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    rows = survey()
    print(json.dumps(rows, ensure_ascii=False, indent=1) if a.json else render(rows))
    # ★ path_fails が在る回だけ落とす ── vague/no_path は在庫であって赤ではない。
    return 1 if any(r["verdict"] == "path_fails" for r in rows) else 0


if __name__ == "__main__":
    raise SystemExit(main())
