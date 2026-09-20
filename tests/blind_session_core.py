# -*- coding: utf-8 -*-
"""盲検の 1 体を**そのまま再生できる形で残し、合格率で採る**（2026-09-20）。

★★ なぜ要るか: 4 体やって、買い手が実際に打った依頼文が**どこにも残っていない**。
  手元の履歴（`~/.ailine/history.jsonl` 17,912 件）を「借方金額」で引いて **0 件** ──
  買い手役は別の `AILINE_HOME` で走っていた。所見は文書に残ったが、**入力が残らない**。
  ★ だから「3 体目が落ちた依頼は、いま通るのか？」に機械で答えられない。
    盲検が 1 回ずつ使い捨てになっていて、積み上がっていなかった。

★★ 直し方は掘り起こしでなく**次の 1 体から構造的に残す**:

    prepare  買い手専用の `AILINE_HOME` を作り、買い手のシェルに貼る行を印字する。
             同時に**いま何を売っているか**（HEAD の sha ＋ 導入済みの版）を刻む。
    freeze   セッション後、その home の履歴から依頼文を取り出し、使った冊ごと固める。
    replay   固めた依頼を**いまの HEAD** に当て直し、**依頼ごとに N 回**振って合格率を出す。
    report   棚（ブルーストロベリー）に積む 1 節を書き出す。

★★ 合格の採り方（2026-09-20・Namakoo 決裁。合格線の文書が正）:
  ・**到達するか**は依頼ごとに **30 回**振って**合格率**で採る
    （Namakoo「一回で合格はかえって疑わしさが残った状態だ。ただ全勝しても
      N+1 回目の動作保証は出来ないから、合格率で判定しよう」）
  ・**嘘の到達**（`✓` が付いたのに中身が違う）には**率を当てない ── 1 度でも出たら不合格**
  ・報告は必ず **(k/N, 下限)** の組で書く。「100%」とは書かない
    （30/30 で言えるのは「真の成功率 90% 以上（片側 95%）」まで）

★ 盲検の作法は変えない ── 買い手は README だけを見る。この道具が触るのは
  「環境の用意」と「終わった後の記録」だけで、買い手の見るものには一切出てこない。

★ なぜ tests/ に在るか: 素の環境の番人は `scripts/` 同士の import を弾く（3 度踏んだ）。
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CORPUS = REPO / "bench" / "blind"

#: ★ 合格率を採る既定の回数（Namakoo 決裁）。30/30 で「真の成功率 90% 以上」が言える最小。
DEFAULT_RUNS = 30


# ── 下限の見積り ────────────────────────────────────────────────────────────
def lower_bound(k: int, n: int, conf: float = 0.95) -> float:
    """観測 k/n から、真の成功率の**片側下限**を返す（Clopper–Pearson）。

    ★★ なぜ生の割合を出さないか: 30/30 は「100%」ではない。言えるのは下限だけで、
      30/30 でも **90% 以上**までしか言えない。「100% 通った」と書くと、
      この repo が何度も踏んだ「観測していないことを主張する」形になる。
    ★ 外の道具に頼らない（素の環境で走る）── 二分探索で不完全ベータの逆を取る。
    """
    if n <= 0:
        return 0.0
    if k <= 0:
        return 0.0
    if k >= n:
        # ★ P(全勝 | p) = p**n = 1-conf を解く（Clopper–Pearson の全勝時の形）
        return (1.0 - conf) ** (1.0 / n)
    # ★ 下限 p は「p のとき k 回以上成功する確率 = 1-conf」を満たす値
    from math import comb

    def at_least(p: float) -> float:
        return sum(comb(n, i) * (p ** i) * ((1 - p) ** (n - i)) for i in range(k, n + 1))

    lo, hi = 0.0, 1.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if at_least(mid) > 1 - conf:
            hi = mid
        else:
            lo = mid
    return lo


def rate_line(k: int, n: int) -> str:
    """報告の 1 行 ── **(k/N, 下限) の組**でしか書かない。"""
    return f"{k}/{n}（真の成功率 {lower_bound(k, n) * 100:.0f}% 以上・片側95%）"


# ── 環境と記録 ──────────────────────────────────────────────────────────────
def _git(*args) -> str:
    r = subprocess.run(["git", *args], cwd=str(REPO), capture_output=True, text=True)
    return (r.stdout or "").strip()


def _env(home: Path | None = None) -> dict:
    e = {**os.environ, "PYTHONPATH": str(REPO / "src")}
    if home is not None:
        e["AILINE_HOME"] = str(home)
    return e


def _version_line() -> str:
    r = subprocess.run([sys.executable, "-m", "ailine", "doctor"], cwd=str(REPO),
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=_env())
    for line in (r.stdout or "").splitlines():
        if "ailine 版" in line:
            return line.strip()
    return "（doctor が版を言わなかった）"


def prepare(name: str, force: bool = False) -> list:
    """買い手専用の `AILINE_HOME` を用意し、貼り付ける行を返す。"""
    d = CORPUS / name
    home = d / "home"
    if home.exists() and not force:
        return [f"？ すでに在ります: {home}（作り直すなら --force）"]
    if home.exists():
        shutil.rmtree(home)
    home.mkdir(parents=True)
    (d / "meta.json").write_bytes(json.dumps({
        "name": name,
        "prepared_at": datetime.now(timezone.utc).isoformat(),
        "head": _git("rev-parse", "HEAD"),
        "version_line": _version_line(),
    }, ensure_ascii=False, indent=1).encode("utf-8"))
    return [
        f"✓ 用意しました: {d}",
        "",
        "★ 買い手役のシェルに、そのまま貼ってください（1 行目が要 ── 4 体目はここが",
        "  継承されておらず『環境を整えてある』が嘘になりました）:",
        "",
        f'    $env:AILINE_HOME = "{home}"',
        # ★★ 2026-09-20: AILINE_TRACE は**フラグではなくパス**（製品は開いて追記する）。
        #   初版は "1" を渡しており、買い手のカレントに `1` という名のファイルを作っていた
        #   ── 「環境を整えてある」がまた嘘になるところだった（4 体目の汚染 2 と同じ形）。
        #   ★ 見つけたのは `check`（この道具自身の環境確認）。設定した、で終わらせない。
        f'    $env:AILINE_TRACE = "{d / "argv.jsonl"}"',
        "",
        "★ 買い手に渡すのは README だけ。src/ tests/ bench/ docs/ は見せません。",
    ]


def _rows(home: Path) -> list:
    p = home / "history.jsonl"
    if not p.exists():
        return []
    return [json.loads(l) for l in p.read_text(encoding="utf-8").splitlines() if l.strip()]


def freeze(name: str) -> list:
    """セッション後、依頼文と**使った冊**を固める（再生が自己完結する）。"""
    d = CORPUS / name
    rows = _rows(d / "home")
    if not rows:
        return [f"？ 履歴が空です: {d / 'home' / 'history.jsonl'}",
                "  ★ 買い手のシェルで AILINE_HOME が効いていなかった可能性があります。"]
    books = d / "books"
    books.mkdir(parents=True, exist_ok=True)
    out = []
    for r in rows:
        src = Path(r.get("book") or "")
        kept = None
        if src.name and src.exists():
            kept = hashlib.sha256(src.read_bytes()).hexdigest()[:12] + src.suffix
            if not (books / kept).exists():
                shutil.copy2(src, books / kept)
        out.append({"ts": r.get("ts"), "task": r.get("task"), "book": kept,
                    "book_name": src.name or None, "path": r.get("path"),
                    "op": r.get("op"), "ok": r.get("ok"), "verdict": r.get("verdict")})
    meta = json.loads((d / "meta.json").read_bytes())
    (d / "requests.json").write_bytes(json.dumps(
        {"head": meta["head"], "version_line": meta.get("version_line"),
         "frozen_at": datetime.now(timezone.utc).isoformat(), "requests": out},
        ensure_ascii=False, indent=1).encode("utf-8"))
    missing = sum(1 for r in out if r["book"] is None)
    lines = [f"✓ 固めました: {len(out)} 件 → {d / 'requests.json'}"]
    if missing:
        lines.append(f"⚠ 冊が見つからなかった依頼が {missing} 件（再生できません）")
    return lines


# ── 再生（合格率） ──────────────────────────────────────────────────────────
_LANDED = re.compile(r"事後条件を確認（操作:([^）]+)）")


def _once(book: Path, task: str, home: Path) -> dict:
    """1 回走らせる ── (exit, 着いた先, 画面の尻)。

    ★ 着いた先は**通った回の画面**が言う（「依頼を『…』と読みました」は断りにしか出ない
      ── 2026-09-20 に測定器を直した。初版は通った回を全部 None と読んでいた）。
    """
    r = subprocess.run([sys.executable, "-m", "ailine", "run", str(book), task,
                        "--copy", "--timeout", "150"],
                       cwd=str(REPO), capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=_env(home))
    out = r.stdout or ""
    m = _LANDED.search(out)
    return {"rc": r.returncode, "landed": m.group(1) if m else None, "tail": out[-200:]}


def replay_one(d: Path, req: dict, runs: int, run_fn=None) -> dict:
    """1 つの依頼を runs 回振り、合格率と**着いた先の顔ぶれ**を返す。"""
    fn = run_fn or _once
    outs = []
    for _ in range(runs):
        with tempfile.TemporaryDirectory() as td:
            home = Path(td) / "home"
            home.mkdir()
            work = Path(td) / (req.get("book_name") or "in.xlsx")
            shutil.copy2(d / "books" / req["book"], work)
            outs.append(fn(work, req["task"], home))
    reached = [o for o in outs if o["rc"] == 0]
    landed = sorted({o["landed"] for o in reached if o["landed"]})
    return {"task": req["task"], "runs": runs, "reached": len(reached),
            "landed": landed,
            # ★★ 2 つ以上の op に着いて**どちらも通った** ── 少なくとも片方は誤り。
            #   これは「嘘の到達」の影が機械で見える唯一の形（全部は見えない）。
            "two_answers": len(landed) > 1,
            "screens": [o["tail"] for o in outs if o["rc"] != 0][:2]}


def verdicts(results: list) -> dict:
    """★ 合格の判定（合格線の文書が正）。

    ★★ 率を当てるのは**到達**だけ。`two_answers`（同じ依頼が 2 つの答えで通った）は
      嘘の到達の影なので、**1 件でも在れば不合格**（率で薄めない）。
    """
    two = [r for r in results if r["two_answers"]]
    total_runs = sum(r["runs"] for r in results)
    total_ok = sum(r["reached"] for r in results)
    return {"requests": len(results), "runs": total_runs, "reached": total_ok,
            "two_answers": [r["task"] for r in two],
            "failed_by_two_answers": bool(two)}


def render(name: str, results: list, v: dict, meta: dict) -> str:
    """画面に出す読み上げ。★ 率は必ず (k/N, 下限) の組で書く。"""
    lines = [f"盲検 {name} ── 再生で採った合格率", ""]
    lines.append(f"  売っていた版: {meta.get('version_line') or '（不明）'}")
    lines.append(f"  固めた時の HEAD: {(meta.get('head') or '')[:8]} → いま {_git('rev-parse', '--short', 'HEAD')}")
    lines.append(f"  依頼 {v['requests']} 件 × {results[0]['runs'] if results else 0} 回 "
                 f"= {v['runs']} 走行")
    lines.append("")
    for r in sorted(results, key=lambda x: x["reached"]):
        mark = "×" if r["two_answers"] else ("○" if r["reached"] == r["runs"] else "△")
        lines.append(f"  {mark} {rate_line(r['reached'], r['runs'])}  「{r['task'][:40]}」")
        if r["two_answers"]:
            lines.append(f"      ★★ 同じ依頼が 2 つの答えで通った: {'・'.join(r['landed'])}")
            lines.append("         ── 少なくとも片方は誤り（嘘の到達の影）")
        elif r["reached"] < r["runs"] and r["screens"]:
            lines.append(f"      断った回の画面: {r['screens'][0].strip()[:110]}")
    lines.append("")
    if v["failed_by_two_answers"]:
        lines.append("★★ 不合格 ── 嘘の到達の影が出ている（率では薄めない・1 件でも不合格）")
    else:
        lines.append("★ 嘘の到達の影は出なかった（ただし**安定して間違っている**分は、"
                     "この器では見えない ── 人が読む側の仕事が残る）")
    return "\n".join(lines)


def shelf_node(name: str, results: list, v: dict, meta: dict) -> str:
    """ブルーストロベリーの棚に積む 1 節（`library-blind/nodes/<name>.md`）。

    ★★ なぜ棚に積むか（Namakoo「ブルーストロベリーに記録を蓄積したい」）: 所見は
      ailine の `docs/` に残るが、それは**この repo の中**にしか無い。棚に置くと
      セッションをまたいで検索・共鳴の対象になり、4 体ぶんが 1 回ずつ使い捨てに
      なっていた状態から抜けられる。
    ★ 棚に書くのは**数と出来事**だけ。生の依頼文と冊は repo 側（`bench/blind/`）に置く
      ── 棚に大きな実体を持ち込まない（索引と実体を分ける既存の作法どおり）。
    """
    worst = sorted(results, key=lambda x: x["reached"])[:5]
    out = [f"# 盲検 {name}", "",
           f"- 採った日: {datetime.now(timezone.utc).date().isoformat()}",
           f"- 売っていた版: {meta.get('version_line') or '（不明）'}",
           f"- 固めた時の HEAD: `{(meta.get('head') or '')[:12]}`",
           f"- 依頼 {v['requests']} 件 × {results[0]['runs'] if results else 0} 回 = {v['runs']} 走行",
           f"- 到達 {rate_line(v['reached'], v['runs'])}（全体）",
           "",
           "★ 率は **(k/N, 下限)** の組でしか書かない ── 全勝でも「100%」ではない",
           "  （30/30 で言えるのは「真の成功率 90% 以上・片側95%」まで）。", ""]
    if v["failed_by_two_answers"]:
        out += ["## ★★ 不合格 ── 嘘の到達の影", "",
                "同じ依頼が **2 つの答えで通った**（少なくとも片方は誤り）:", ""]
        out += [f"- 「{t}」" for t in v["two_answers"]]
        out.append("")
    out += ["## 低い順に 5 件", ""]
    out += [f"- {rate_line(r['reached'], r['runs'])} 「{r['task'][:50]}」" for r in worst]
    out += ["", "## 実体の在り処", "",
            f"- 依頼と冊: `{CORPUS / name}`",
            f"- 合格線: `{REPO / 'docs' / 'FROZEN-20260917-盲検の合格線.md'}`",
            "", "関連: [[blind-how-we-score]]", ""]
    return "\n".join(out)


# ── 環境確認 ────────────────────────────────────────────────────────────────
#
# ★★ なぜ在るか（4 体目の汚染 1・2 ── どちらもこちらの落ち度）:
#   ・入っていた `ailine` が古く、買い手は「README に在るコマンドが無い」で **30 分**溶かした
#   ・`AILINE_TRACE` が買い手のシェルに継承されておらず、「環境を整えてある」が嘘になった
#   ★ どちらも「設定したつもり」で止まっていた。**効いていることを測る**装置にする
#     （この repo の古い線: 設定≠動く／直した≠反映／調べた≠確かめた）。


def readme_commands() -> set:
    """README が買い手に打たせるサブコマンド（`ailine <語>`）。"""
    text = (REPO / "README.md").read_bytes().decode("utf-8")
    return set(re.findall(r"ailine ([a-z][a-z-]+)", text))


def parser_commands() -> set:
    """製品が実際に持つサブコマンド ── **argparse から導く**（手で並べない）。"""
    sys.path.insert(0, str(REPO / "src"))
    import ailine as _a
    import argparse as _ap
    out = set()
    for act in _a.build_parser()._actions:
        if isinstance(act, _ap._SubParsersAction):
            out |= set(act.choices)
    return out


def _count_lines(p: Path) -> int:
    if not p.exists():
        return 0
    return len([l for l in p.read_text(encoding="utf-8").splitlines() if l.strip()])


def check(name: str, run_fn=None) -> list:
    """買い手に渡す前の環境確認 ── 各項は **(鍵, ok, 一行)** を返す。

    ★★ 鍵を持たせる理由（2026-09-20）: 初版は (ok, 一行) だけで、番人は一行の**文面**から
      鍵を切り出していた。赤い回は文面が変わるので、番人が拾えなくなる ──
      今週 5 件目の「番人を字面で書いた」形。**鍵は機械が持つ**（文面は人向け）。

    ★★ 測る順は「渡す前に潰せる順」── 版 → README → 隔離 → 控え → 前提。
    ★ 3 番目と 4 番目は**実際に 1 回走らせて**確かめる。設定を読んで安心しない。
    """
    d = CORPUS / name
    home = d / "home"
    trace = d / "argv.jsonl"
    rows = []

    # ① 版 ── 作業木と導入済みが揃っているか（4 体目はここで 30 分溶けた）
    v = _version_line()
    rows.append(("版", v.startswith("✓"), f"版: {v[:110]}"))

    # ② README が言うコマンドが**実際に在る**か（買い手が最初に触るのは README だけ）
    missing = sorted(readme_commands() - parser_commands())
    rows.append(("README", not missing,
                 "README のコマンド: 全部在る" if not missing
                 else f"★ README に在って製品に無い: {missing}"))

    if not home.exists():
        rows.append(("未用意", False, f"★ 先に prepare を走らせてください（{home} が無い）"))
        return rows

    # ③④ 実際に 1 回走らせて、隔離と控えが**効いている**ことを見る
    #   ★ 買い手が使う**その設定のまま**打つ（別の器で確かめても、その設定の証拠にならない）。
    #   ★★ ただし測定が対象を汚さないよう、**終わったら元のバイトに戻す** ──
    #     戻さないと、この試し打ちが冊の無い依頼として corpus に混ざる（実測で気づいた）。
    hist = home / "history.jsonl"
    default_home = Path.home() / ".ailine" / "history.jsonl"
    # ★★ 2026-09-20: 後始末は**home の中身ぜんぶ**を見る（2 ファイルだけ戻していた初版は
    #   `notice_v2_shown` を残し、**買い手が見るはずの初回の告知を試し打ちが消費して**いた）。
    #   ★ 「測定が対象を汚さない」を 2 度続けて踏んだ ── 対象は狭く数えない。
    keep = {f: f.read_bytes() for f in home.rglob("*") if f.is_file()}
    keep[trace] = trace.read_bytes() if trace.exists() else None
    before_here, before_default = _count_lines(hist), _count_lines(default_home)
    before_trace = _count_lines(trace)
    try:
        ok_run, tail = (run_fn or _probe)(home, trace)
        after_here, after_default = _count_lines(hist), _count_lines(default_home)
        after_trace = _count_lines(trace)
    finally:
        for f, b in keep.items():
            if b is None:
                f.unlink(missing_ok=True)
            else:
                f.write_bytes(b)
        # ★ 試し打ちが**新しく作った**ファイルも消す（差分でなく現物を突き合わせる）
        for f in list(home.rglob("*")):
            if f.is_file() and f not in keep:
                f.unlink(missing_ok=True)

    rows.append(("試し打ち", ok_run,
                 f"試し打ち: {'通った' if ok_run else '★ 落ちた ── ' + tail[:90]}"))
    # ★★ 両側から見る ── 「こちらに増えた」だけでは隔離の証拠にならない。
    #   **本体の履歴が増えていない**ことまで確かめて、はじめて隔離が効いている。
    rows.append(("AILINE_HOME", after_here > before_here and after_default == before_default,
                 f"AILINE_HOME: この回 +{after_here - before_here} / "
                 f"本体 +{after_default - before_default}"
                 + ("" if after_default == before_default else "  ★ 本体に漏れている")))
    rows.append(("AILINE_TRACE", after_trace > before_trace,
                 f"AILINE_TRACE: 控え +{after_trace - before_trace} 行（{trace.name}）"))
    rows.append(("後始末",
                 _count_lines(hist) == before_here and _count_lines(trace) == before_trace,
                 "後始末: 試し打ちの跡を消した（corpus を汚さない）"))
    return rows


def _probe(home: Path, trace: Path) -> tuple:
    """環境確認のための 1 回（★ 製品を素の形で打つ ── 特別扱いしない）。"""
    import openpyxl
    with tempfile.TemporaryDirectory() as td:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "表"
        ws.append(["品名", "金額"])
        ws.append(["机", 12000])
        book = Path(td) / "確認.xlsx"
        wb.save(book)
        wb.close()
        env = _env(home)
        env["AILINE_TRACE"] = str(trace)
        r = subprocess.run([sys.executable, "-m", "ailine", "run", str(book),
                            "金額の大きい順に並べ替えて", "--copy", "--sheet", "表",
                            "--timeout", "150"],
                           cwd=str(REPO), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", env=env)
    return r.returncode == 0, (r.stdout or "")[-200:]


def render_check(name: str, rows: list) -> str:
    lines = [f"盲検 {name} ── 買い手に渡す前の環境確認", ""]
    for _key, ok, text in rows:
        lines.append(f"  {'✓' if ok else '×'} {text}")
    lines.append("")
    bad = [t for _k, ok, t in rows if not ok]
    lines.append("★ 渡してよい" if not bad
                 else f"★★ まだ渡さない ── {len(bad)} 件が効いていない")
    return "\n".join(lines)
