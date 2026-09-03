# -*- coding: utf-8 -*-
"""`verify_dsl_args` の (入力, 出力) を全テストから記録する ── 組み替えの前後で比べる基準線。

★ なぜ在るか（2026-09-04）: README は長く「単一ファイルを割るべきだが、**挙動不変を
  確かめる番人**を用意できていない」と書いてきた。事後条件 45 関数の移動は
  `inspect.getsource` の一致で確かめられたが、`verify_dsl_args`（1,735 行・op 分岐 30 個・
  5 本の入れ子チェーン）は**移動でなく構造の組み替え**になるので、同じ手が使えない。

★ 使える手が測定で見えた: テストを 1 回走らせると **29 op ぶん 633 回**この関数を通る。
  その (入力, 出力) を丸ごと凍結すれば、組み替えの前後で **1 件でも動いたら赤**にできる。

    # 基準線を取る（組み替える前に）
    AILINE_GOLDEN_OUT=bench/verify_golden.json PYTHONPATH="src;tests" \
      python -m pytest tests/ -q -p verify_golden_probe

    # 組み替えた後に、同じものを取って比べる
    python bench/verify_golden_probe.py --compare bench/verify_golden.json new.json

★★ 記録の設計（何を残し、何を残さないか）:
  ・入力  op / args / task / target_sheet / vocab の有無 / book_meta の**要約**
    ★ book_meta は丸ごとだと巨大なので、シート名・見出し・パスの有無だけに畳む。
      畳んだぶん「同じ入力に見えて実は違う」が起きうるので、★ その可能性を認めた上で使う
  ・出力  ok / resolved / inferred / err を**全部**
    ★ ここは畳まない。組み替えで変わるのは出力側だから
  ・順序は記録しない（テストの実行順に依存させない ── 並べ替えて比較する）
"""
from __future__ import annotations

import argparse
import atexit
import functools
import hashlib
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(os.environ.get("AILINE_GOLDEN_OUT", "verify_golden.json"))
_rec: list = []


def _plain(v, depth=0):
    """JSON に載る形へ。★ 深追いしない（set は並べ替えて list に）。"""
    if depth > 6:
        return "…"
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    if isinstance(v, set):
        return ["<set>"] + sorted(_plain(x, depth + 1) for x in v)
    if isinstance(v, (list, tuple)):
        return [_plain(x, depth + 1) for x in v]
    if isinstance(v, dict):
        return {str(k): _plain(x, depth + 1) for k, x in sorted(v.items(), key=lambda kv: str(kv[0]))}
    return f"<{type(v).__name__}>"


def _meta_digest(book_meta) -> dict:
    """book_meta を要約する（丸ごとは巨大なので）。"""
    if not isinstance(book_meta, dict):
        return {"_": f"<{type(book_meta).__name__}>"}
    heads = book_meta.get("headers") or {}
    return {
        "sheets": _plain(book_meta.get("sheets")),
        "headers": _plain(heads),
        "sheet_source": _plain(book_meta.get("_sheet_source")),
        "has_path": bool(book_meta.get("path")),
    }


def install():
    import ailine
    orig = ailine.verify_dsl_args

    # ★ functools.wraps で**署名と名前を保つ**（2026-09-04）:
    #   これが無いと `inspect.signature(ailine.verify_dsl_args)` が変わり、
    #   公開面の凍結（tests/test_public_surface_is_frozen.py）が必ず赤くなる。
    #   ★ 基準線を取るたびに 1 件赤くなる状態は、**本物の赤を見落とす原因**になる。
    @functools.wraps(orig)
    def probe(op, args, book_meta, task="", vocab=None, target_sheet=None):
        r = orig(op, args, book_meta, task=task, vocab=vocab, target_sheet=target_sheet)
        try:
            inp = {"op": op, "args": _plain(args), "task": task,
                   "target_sheet": target_sheet, "has_vocab": vocab is not None,
                   "book": _meta_digest(book_meta)}
            out = {"ok": bool(r[0]), "resolved": _plain(r[1]),
                   "inferred": _plain(r[2]), "err": _plain(r[3])}
            key = hashlib.sha1(
                json.dumps(inp, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:16]
            _rec.append({"key": key, "in": inp, "out": out})
        except Exception as e:                      # ★ 記録に失敗しても本体は止めない
            _rec.append({"key": "?", "error": f"{type(e).__name__}: {e}"})
        return r

    ailine.verify_dsl_args = probe


def _dump():
    if not _rec:
        return
    rows = sorted(_rec, key=lambda r: (r.get("key", ""),
                                        json.dumps(r.get("out", {}), ensure_ascii=False,
                                                   sort_keys=True)))
    OUT.write_text(json.dumps({"count": len(rows), "rows": rows},
                               ensure_ascii=False, indent=1), encoding="utf-8")


def compare(a: Path, b: Path) -> int:
    """2 つの基準線を比べる。★ 1 件でも出力が動いたら赤。"""
    da, db = (json.loads(p.read_text(encoding="utf-8")) for p in (a, b))
    ia = {r["key"]: r for r in da["rows"] if "out" in r}
    ib = {r["key"]: r for r in db["rows"] if "out" in r}
    print(f"前 {da['count']} 件 / 後 {db['count']} 件")
    gone = sorted(set(ia) - set(ib))
    added = sorted(set(ib) - set(ia))
    changed = [k for k in set(ia) & set(ib) if ia[k]["out"] != ib[k]["out"]]
    if gone:
        print(f"★ 通らなくなった入力: {len(gone)} 件")
        for k in gone[:5]:
            print(f"    {ia[k]['in']['op']}  {str(ia[k]['in']['task'])[:52]}")
    if added:
        print(f"（新しく通った入力: {len(added)} 件 ── 検体を足したなら正当）")
    if changed:
        print(f"★★ 出力が変わった: {len(changed)} 件")
        for k in changed[:5]:
            print(f"    {ia[k]['in']['op']}  {str(ia[k]['in']['task'])[:44]}")
            print(f"      前 ok={ia[k]['out']['ok']} err={str(ia[k]['out']['err'])[:44]}")
            print(f"      後 ok={ib[k]['out']['ok']} err={str(ib[k]['out']['err'])[:44]}")
    if not gone and not changed:
        print("✓ 出力は 1 件も変わっていない（挙動不変）")
        return 0
    return 1


# --- pytest プラグインとして読み込まれた時 -----------------------------------------------
if "pytest" in sys.modules or os.environ.get("AILINE_GOLDEN_OUT"):
    sys.path.insert(0, str(ROOT / "src"))
    try:
        install()
        atexit.register(_dump)
    except Exception:
        pass


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--compare", nargs=2, metavar=("BEFORE", "AFTER"))
    a = ap.parse_args(argv)
    if a.compare:
        return compare(Path(a.compare[0]), Path(a.compare[1]))
    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
