# -*- coding: utf-8 -*-
"""盤の前後を**マス単位**で突き合わせる（2026-09-17）。

★ なぜ在るか（Namakoo）:
  「変更により別の項目に○がついたり、退行により落ちたりする可能性の見落とし」。
  直した列だけ見て済ませると、隣の列が動いたのに気づかない。
  golden で「増えた行は 2 種類だけ」と機械で確かめたのと同じことを、盤にもやる。

★ 見るのは 4 つ:
    ① 列の増減
    ② 列の**立場**の変化（導出あり / 番人あり / 導けない / ★ 未調査 / ★ 無防備）
    ③ **マスの増減**（どの op がどの名簿に入った・出た）
    ④ 宣言と実測の食い違いの増減

★ 使い方:
      python scripts/board_diff.py --save before.json     # 変える前に基準線を取る
      （変更）
      python scripts/board_diff.py --against before.json  # 何が動いたかを全部出す
"""
import argparse
import json
import pathlib
import subprocess
import sys

# ★ 隣の実物を呼ぶ（写しを作ると 2 つになり、2026-09-16 にそれで盤が嘘をついた）
BOARD = pathlib.Path(__file__).with_name("wiring_board.py")
STANCE_JA = {"derived": "導出あり", "watched": "番人あり", "explained": "導けない",
             "unstudied": "★ 未調査", "bare": "★ 無防備"}


def snapshot() -> dict:
    r = subprocess.run([sys.executable, str(BOARD), "--json"],
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        raise SystemExit("盤の生成に失敗:\n" + r.stderr[-600:])
    return json.loads(r.stdout)


def _cells(d: dict) -> dict:
    """(op, 列) → 在籍。マス単位で比べるための形。"""
    out = {}
    for c in d["columns"]:
        for op in c["declared"]:
            out[(op, c["name"])] = True
    return out


def compare(before: dict, after: dict) -> int:
    moved = 0
    b_cols = {c["name"]: c for c in before["columns"]}
    a_cols = {c["name"]: c for c in after["columns"]}

    print("① 列の増減")
    added, removed = sorted(set(a_cols) - set(b_cols)), sorted(set(b_cols) - set(a_cols))
    for n in added:
        print(f"   ＋ {n}")
    for n in removed:
        print(f"   － {n}")
    moved += len(added) + len(removed)
    if not (added or removed):
        print("   変化なし")

    print("\n② 立場の変化")
    n_stance = 0
    for n in sorted(set(a_cols) & set(b_cols)):
        b, a = b_cols[n]["stance"], a_cols[n]["stance"]
        if b != a:
            arrow = "→"
            print(f"   {n:32} {STANCE_JA[b]} {arrow} {STANCE_JA[a]}")
            n_stance += 1
    moved += n_stance
    print("   変化なし" if not n_stance else f"   計 {n_stance} 列")

    print("\n③ マスの増減（★ ここが見落としやすい ── 直した列の隣が動く）")
    bc, ac = _cells(before), _cells(after)
    gained, lost = sorted(set(ac) - set(bc)), sorted(set(bc) - set(ac))
    for op, col in gained:
        print(f"   ＋ {col:32} に {op}")
    for op, col in lost:
        print(f"   － {col:32} から {op}")
    moved += len(gained) + len(lost)
    if not (gained or lost):
        print("   変化なし")

    print("\n④ 宣言と実測の食い違い")

    def mis(d):
        return {c["name"] for c in d["columns"]
                if c["derived"] is not None and c["declared"] != c["derived"]}
    bm, am = mis(before), mis(after)
    for n in sorted(am - bm):
        print(f"   ＋ {n}（新しくずれた）")
    for n in sorted(bm - am):
        print(f"   － {n}（揃った）")
    moved += len(am ^ bm)
    if not (am ^ bm):
        print("   変化なし")

    print(f"\n動いた点: {moved}")
    if moved == 0:
        print("★ 何も動いていない ── 盤に出ない変更だったか、変更が効いていない")
    return moved


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--save", help="いまの盤をここに保存（基準線）")
    ap.add_argument("--against", help="この基準線と突き合わせる")
    a = ap.parse_args(argv)
    now = snapshot()
    if a.save:
        pathlib.Path(a.save).write_bytes(
            json.dumps(now, ensure_ascii=False, indent=1).encode("utf-8"))
        cols = now["columns"]
        print(f"基準線を保存: {a.save}")
        print(f"  {now['commit'][:48]}")
        print(f"  列 {len(cols)} / マス {sum(len(c['declared']) for c in cols)}")
        return 0
    if not a.against:
        ap.error("--save か --against のどちらかが要る")
    before = json.loads(pathlib.Path(a.against).read_bytes().decode("utf-8"))
    print(f"基準線: {before['commit'][:48]}")
    print(f"いま　: {now['commit'][:48]}\n")
    compare(before, now)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
