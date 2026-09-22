"""盲検の記録を数える芯（2026-09-22）。

★ 様式は docs/盲検の記録様式.md で固定した。ここはそれを読む側。
★ 番人（test_the_blind_record_can_be_counted.py）と、率を出す道具の両方が使う。
"""
import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DOCS = REPO / "docs"

#: 様式を固定した日。★ これ**以降**の記録だけを縛る ──
#: 過去の記録を遡って様式に合わせたら、それは記録でなくなる。
SCHEMA_FROM = "20260922"

KINDS = ("欠陥", "摩擦", "要望")
REPRO = ("再現", "未再現", "引用で明白")

_ROW = re.compile(r"^\|\s*([①-⑳]+)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|"
                  r"\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*$")


def blind_records() -> list:
    """盲検の記録ファイル（新しい順）。"""
    return sorted(DOCS.glob("DESIGN-*買い手役*.md"), reverse=True)


def records_under_the_schema() -> list:
    """様式を固定した日**以降**に書かれた記録だけ。"""
    out = []
    for p in blind_records():
        m = re.search(r"DESIGN-(\d{8})", p.name)
        if m and m.group(1) >= SCHEMA_FROM:
            out.append(p)
    return out


def parse_findings(path: Path) -> list:
    """様式どおりの行を [{mark, kind, text, repro, fatal}] で返す。"""
    out = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = _ROW.match(ln)
        if not m:
            continue
        mark, kind, text, repro, fatal = (g.strip() for g in m.groups())
        if kind not in KINDS:          # ★ 見出し行や別の表を拾わない
            continue
        out.append({"mark": mark, "kind": kind, "text": text,
                    "repro": repro, "fatal": fatal})
    return out


def required_fields(path: Path) -> dict:
    """`分母:` `役柄:` `渡したもの:` が書かれているか。"""
    text = path.read_text(encoding="utf-8")
    return {k: bool(re.search(rf"^\s*[-*]?\s*`?{k}:?`?\s*\S", text, re.M))
            for k in ("分母", "役柄", "渡したもの")}


def rate(findings: list) -> dict:
    """欠陥の数と、外の目の誤報率。"""
    defects = [f for f in findings if f["kind"] == "欠陥"]
    not_repro = [f for f in findings if f["repro"] == "未再現"]
    return {
        "所見": len(findings),
        "欠陥": len(defects),
        "再現した欠陥": len([f for f in defects if f["repro"] != "未再現"]),
        "未再現": len(not_repro),
        "誤報率": (len(not_repro) / len(findings)) if findings else None,
        "致命": len([f for f in findings if f["fatal"] == "致命"]),
    }
