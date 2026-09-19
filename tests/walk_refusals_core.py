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


def _run(argv: list, plan) -> tuple:
    """製品を 1 回走らせ、(exit, 画面) を返す。plan を渡せば翻訳をそれに固定する。"""
    buf = io.StringIO()
    real = ailine.translate_task
    if plan is not None:
        ailine.translate_task = lambda *a, **k: {"plan": plan}
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                rc = ailine.main(argv)
            except SystemExit as e:
                rc = e.code
    finally:
        ailine.translate_task = real
    return rc, buf.getvalue()


def _example_in(text: str) -> str | None:
    """断り文が見せている「通る書き方」を取り出す（『…』か「…」の中で、依頼文らしいもの）。"""
    import re
    for m in re.finditer(r"[『「]([^』」]{6,60})[』」]", text):
        s = m.group(1)
        if any(w in s for w in ("抜き出", "並べ替え", "追加", "にして", "消して", "入れて")):
            return s
    return None


def walk_one(key: str, entry: dict, root: Path) -> dict:
    """1 件の断りを出させ、示された道を歩く。戻り値に verdict を入れる（台帳は読まない）。"""
    w = entry.get("walk")
    if not w:
        return {"key": key, "verdict": "未記入", "detail": "walk 欄が無い"}
    folder = _make_folder(root)
    argv = ["run", str(folder), w["task"], "--out", str(root / "結果.xlsx")]
    rc, out = _run(argv, w.get("plan"))
    if rc == 0:
        return {"key": key, "verdict": "引き金が引けない",
                "detail": "断りが出ずに到達した（検体が古い）", "screen": out[-200:]}

    path = w.get("path") or {}
    kind = path.get("kind")
    if kind == "none":
        # ★ 道を示していない ── 理由が粗いのか、本当に道が無いのかは人が仕分ける。
        return {"key": key, "verdict": "vague" if w.get("why") else "no_path",
                "detail": w.get("why") or "示す道が無い", "screen": out.strip()[-160:]}

    if kind == "single_book":
        one = folder / next(iter(BOOKS))
        rc2, out2 = _run(["run", str(one), w["task"], "--copy"], w.get("plan"))
    elif kind == "example":
        task2 = path.get("task") or _example_in(out)
        if not task2:
            return {"key": key, "verdict": "vague",
                    "detail": "例を示していると数えられているが、歩ける文が取り出せない",
                    "screen": out.strip()[-160:]}
        # ★★ 2026-09-19（最初の実行で踏んだ測定器の穴）: ここで plan=None にすると
        #   **本物の LLM へ投げる**ことになり、ollama に届かない回は FREEFORM が返って
        #   「道が通らない」と誤判定した（盤を手で回した時はたまたま通り、番人では落ちた
        #   ── 測定器が揺れていた）。
        #   ★ 歩く時も計画を固定する。確かめたいのは「**この例文で、示された操作に届くか**」
        #     であって、LLM がその例文をどう読むかではない（それは別の測定）。
        rc2, out2 = _run(["run", str(folder), task2, "--out", str(root / "結果2.xlsx")],
                          path.get("plan") or [{"op": "EXTRACT",
                                                "args": {"col": "金額", "cmp": "gte", "value": 100}}])
    else:
        return {"key": key, "verdict": "未記入", "detail": f"歩き方が不明: {kind}"}

    if rc2 == 0:
        return {"key": key, "verdict": "walked", "detail": f"道を歩いて到達（{kind}）"}
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


ORDER = ["path_fails", "vague", "no_path", "未記入", "引き金が引けない", "walked", "by_design"]


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
