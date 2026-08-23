"""B1 忠実度ベースライン（★ W8b 項目7）。

bench/realworld/*.xlsx の各検体を実際に LibreOffice の正規化パス
（ailine.normalize_book・何もしない空マクロで一度開いて保存するだけ）に通し、
「LO 往復『だけ』で失われるもの」を ailine.check_round_trip_fidelity で測る。

これは cmd_run --inplace の往復忠実度ゲートが本番で検出する対象そのものであり、
「既定を .out から原本直接へ反転してよいか」を判断する材料（W8b 第二コミット向け）。

opt-in・bench 層専用（resilience_check.py と同じ理由で cmd_run の既定パイプラインには
混ぜない — 実行のたびに LO を何度も起動するのは重い）。

実行: python bench/fidelity_baseline.py
出力: 標準出力の表 + bench/realworld/FIDELITY.md
"""
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).parent
REALWORLD = HERE / "realworld"
WORK = HERE / "fidelity_baseline_work"
AILINE_DIR = HERE.parent
sys.path.insert(0, str(AILINE_DIR / "src"))
import ailine  # noqa: E402

# ★ make_specimens.py が生成する検体。shapes.xlsx は openpyxl に高水準 API が無く
#   zip 後処理で作った合成品のため、そもそも作れなかった/LO で開けなかった場合は
#   表に「作れない」「開けない」と正直に記録する（無理に数字を埋めない）。
SPECIMENS = [
    ("A title_rows", "title_rows.xlsx"),
    ("B large(3000行)", "large.xlsx"),
    ("C formulas", "formulas.xlsx"),
    ("D merged_head", "merged_head.xlsx"),
    ("E cf(条件付き書式)", "cf.xlsx"),
    ("F datavalidation(入力規則)", "datavalidation.xlsx"),
    ("G shapes(図形)", "shapes.xlsx"),
    ("H rows1500(1500行)", "rows1500.xlsx"),
]


def measure_one(label: str, filename: str) -> dict:
    src = REALWORLD / filename
    if not src.exists():
        return {"label": label, "file": filename, "status": "検体が無い",
                 "detail": "bench/realworld/make_specimens.py を先に実行して", "fidelity": None}

    work = WORK / Path(filename).stem
    if work.exists():
        shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    copy = work / filename
    shutil.copy2(src, copy)

    try:
        normalized = ailine.normalize_book(copy, work, timeout=180.0)
    except SystemExit as e:
        return {"label": label, "file": filename, "status": "LibreOffice で開けない",
                 "detail": str(e), "fidelity": None}
    except Exception as e:
        return {"label": label, "file": filename, "status": "正規化で例外",
                 "detail": f"{type(e).__name__}: {e}", "fidelity": None}

    fidelity = ailine.check_round_trip_fidelity(src, normalized)
    status = "喪失あり" if fidelity["lost"] else "喪失なし"
    detail = "・".join(f"{it['label']} {it['count']}件" for it in fidelity["items"]) or "-"
    return {"label": label, "file": filename, "status": status, "detail": detail, "fidelity": fidelity}


def format_markdown(rows: list) -> str:
    lines = [
        "# B1 忠実度ベースライン（★ W8b 項目7）",
        "",
        "`bench/fidelity_baseline.py` の実測結果。各検体を `ailine.normalize_book`"
        "（LibreOffice で一度・何もしない空マクロで開いて保存するだけ）に通し、"
        "`check_round_trip_fidelity` で LO 往復だけによる喪失を見る。",
        "",
        "| 検体 | ファイル | 結果 | 内訳 |",
        "|---|---|---|---|",
    ]
    for r in rows:
        lines.append(f"| {r['label']} | {r['file']} | {r['status']} | {r['detail']} |")
    lines.append("")
    lost_count = sum(1 for r in rows if r["fidelity"] and r["fidelity"]["lost"])
    unreadable = sum(1 for r in rows if r["fidelity"] is None)
    lines.append(f"喪失あり: {lost_count} / 測定できた検体 {sum(1 for r in rows if r['fidelity'] is not None)}"
                 f"（開けなかった/作れなかった検体: {unreadable}）")
    return "\n".join(lines)


def main() -> int:
    WORK.mkdir(exist_ok=True)
    rows = [measure_one(label, filename) for label, filename in SPECIMENS]
    print(f"{'検体':<26} {'結果':<14} 内訳")
    for r in rows:
        print(f"{r['label']:<26} {r['status']:<14} {r['detail']}")
    md = format_markdown(rows)
    out_path = REALWORLD / "FIDELITY.md"
    # ★ newline="" で明示（既定の write_text は Windows で LF→CRLF に化ける・
    #   repo は LF 統一のためこれを避ける）。
    out_path.write_text(md, encoding="utf-8", newline="\n")
    print(f"\n書き出し: {out_path}")
    shutil.rmtree(WORK, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
