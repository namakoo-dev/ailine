# -*- coding: utf-8 -*-
"""盲検 1 回ぶんを**凍らせて・流し直して・採る** ── 中身（2026-09-18）。

★★ なぜ tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  「requirements-dev.txt に無い import」を全部止める。scripts/ に置いて番人から
  import すると、**自分の repo の道具なのに弾かれる** ── 2026-09-16 に refresh_records で
  踏み、wiring_board では先回りして避けたのに、**2026-09-18 にまた踏んだ**（3 度目）。
  だから中身はここ、scripts/blind_run.py は薄い入口だけにする。


★ 出所: docs/FROZEN-20260917-盲検の合格線.md「4 体目を回す前にやること ①分母の確保」。
★ Namakoo:「どんな指示に対しても到達するか、適切な断りを出すこと」── その**分母**を
  機械で閉じるための道具。合否そのものは決めない（下の「採らないもの」を読むこと）。

    python scripts/blind_run.py freeze  <回> <作業フォルダ>   # 回が終わった直後に凍らせる
    python scripts/blind_run.py replay  <回>                  # 凍らせた並びを流し直す
    python scripts/blind_run.py floor   <回>                  # ★ 揺れの床を測る（2 回流して差を見る）
    python scripts/blind_run.py sheet   <回>                  # 採点表（機械で決まる 5 条だけ）

★ 回し方（盲検の側）:
    AILINE_HOME=<回のフォルダ>/home  AILINE_TRACE=<回のフォルダ>/trace.jsonl
  を立てて買い手役に渡す。これで**打った全コマンド**と**その回だけの履歴**が集まる。
  ★ history に `ops` / `doctor` を足す案は採らなかった ── history は買い手の月次の証跡で、
    読むだけのコマンドが並ぶと雑音になる（2026-09-18 ⑨の決裁と同じ線）。

★★ 採らないもの（混ぜると「機械が合格と言った」が独り歩きする）:
  9 条のうち機械で決まるのは **5 つだけ** ── 成果が消えない／戻せる／通る道を示す／
  終了コードが成功でない／1 バイトも書かない。
  残り 4 つ（中身が正しい・頼んでいないものを変えない・本当の理由を言う・文言が正確）は
  **人が読む**。しかもその人も盲検にする（判定者の独立性だけで数値が 2.5 倍動いた実測が在る）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
RUNS = REPO / "tests" / "blind_runs"


def _digests(folder: Path) -> dict:
    """**利用者の冊**の指紋（相対パス → sha）。

    ★ 「成果が消えない」「1 バイトも書かない」はここでしか測れない ──
      画面の言葉ではなく**実体**で見る。

    ★★ 2026-09-18: 初版はフォルダ配下を**全部**数えていて、作った当日に嘘を 2 件出した:
      ・断りの回に「断ったのに書いた」── 数えていたのは `home/history.jsonl`。
        合格線の「1 バイトも書かない」は**利用者の冊**の話で、道具の台帳ではない。
        しかも履歴に断りを残すのは 2026-09-18 に足したばかりの**正しい**振る舞い。
      ・揺れの床に 1 件 ── バックアップのファイル名に**時刻**が入るので、
        2 回流せば必ず別名になる。揺れではなく時計だった。
      ★ 測定器が自分の粗さで欠陥を作る形（この repo が何度も踏んでいる）。
        見る範囲を**利用者から見えるファイル**に限る。
    """
    out = {}
    for p in sorted(folder.rglob("*")):
        if not p.is_file():
            continue
        rel = str(p.relative_to(folder)).replace("\\", "/")
        if rel.startswith("home/"):
            continue      # ★ 道具の台帳・バックアップ・別名簿（利用者の成果物ではない）
        out[rel] = hashlib.sha256(p.read_bytes()).hexdigest()[:16]
    return out


def _strip_local_paths(text: str, workdir: Path) -> str:
    r"""凍らせる前に、**その人の実パス**を作業フォルダからの相対へ畳む。

    ★★ 2026-09-18: これが無くて、自己検査を commit しようとしたら機密語の番人が
      6 件で止めた ── `C:\Users\...` が**公開 repo に載る**ところだった。
      ★ 自己検査だけの問題ではない: 本物の盲検を凍らせる時、trace/history には
        **買い手の環境のパス**（氏名やフォルダ構成を含みうる）がそのまま入る。
        凍らせる物は repo に置く物なので、ここで畳むのが唯一の正しい場所。
    ★ 畳むのは**作業フォルダの下だけ**。それ以外の絶対パスが残っていたら、
      それは想定外なので**番人が commit を止める**（黙って消さない）。
    """
    raw = str(workdir)
    # ★ 長い形から先に畳む ── JSON の中では区切りが二重（\\）で入るので、
    #   素の形を先に置換すると二重の側が半端に残る（実測で踏んだ）。
    for form in sorted({raw.replace("\\", "\\\\"), raw, raw.replace("\\", "/")},
                       key=len, reverse=True):
        text = text.replace(form, ".")
    return text


def _run_dir(n: str) -> Path:
    return RUNS / n


# --- freeze ---------------------------------------------------------------------------

def cmd_freeze(n: str, workdir: str) -> int:
    """回が終わった直後に、初期の冊・打った並び・その回の履歴を凍らせる。

    ★ 冊は**買い手が触る前**の物が要る。触った後の冊では流し直せない
      （原本直接の run は冊を変えていくので、途中の状態は再現できない）。
      ★ だから盲検を出す前に books/ へ写しを取っておくこと ── この道具は
        「もう凍っているか」を確かめ、無ければ**名指しで止まる**（黙って進めない）。
    """
    d, w = _run_dir(n), Path(workdir).resolve()
    books = d / "books"
    if not books.is_dir() or not any(books.iterdir()):
        print(f"× {books} に**初期の冊**が在りません。"
              "盲検を出す前に、渡した冊の写しをここへ置いてください"
              "（触られた後の冊では流し直せません）。")
        return 3
    for name, src in (("trace.jsonl", w / "trace.jsonl"),
                      ("history.jsonl", w / "home" / "history.jsonl")):
        if not src.is_file():
            print(f"× {src} が在りません（AILINE_TRACE / AILINE_HOME を立てて回しましたか）。")
            return 3
        (d / name).write_text(_strip_local_paths(
            src.read_text(encoding="utf-8"), w), encoding="utf-8", newline="\n")
    n_cmd = len([1 for _ in (d / "trace.jsonl").read_text(encoding="utf-8").splitlines() if _])
    n_run = len([1 for _ in (d / "history.jsonl").read_text(encoding="utf-8").splitlines() if _])
    print(f"凍らせた: {d}")
    print(f"  打った全コマンド {n_cmd} / うち run 系として履歴に残ったもの {n_run}")
    print(f"  ★ 差の {n_cmd - n_run} 件は ops / doctor / scan / verify など"
          "（history には残らない ── これが分母から落ちていた）")
    return 0


# --- replay ---------------------------------------------------------------------------

def _replay_once(n: str, tag: str) -> Path:
    """凍らせた並びを、冊の**新しい写し**の上で頭から流す。

    ★ 1 つずつ再現しない ── 原本直接の run は冊を変えていくので、
      並びを崩すと「その時の冊」が再現できない。
    """
    d = _run_dir(n)
    work = d / "_replay" / tag
    if work.exists():
        shutil.rmtree(work)
    (work / "home").mkdir(parents=True)
    for p in (d / "books").iterdir():
        if p.is_file():
            shutil.copy2(p, work / p.name)
    env = dict(os.environ, AILINE_HOME=str(work / "home"),
               PYTHONPATH=str(REPO / "src"))
    env.pop("AILINE_TRACE", None)          # ★ 流し直しは控えない（控えの控えを作らない）
    steps = []
    for line in (d / "trace.jsonl").read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        argv = json.loads(line)["argv"]
        # ★ 冊のパスは**この回の写し**へ差し替える（凍らせた時の絶対パスは他所の物）
        argv = [str(work / Path(x).name) if (Path(x).suffix in (".xlsx", ".csv")
                                              and Path(x).name in {q.name for q in work.iterdir()})
                else x for x in argv]
        before = _digests(work)
        r = subprocess.run([sys.executable, "-m", "ailine", *argv], cwd=str(work),
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env)
        after = _digests(work)
        steps.append({
            "argv": argv, "exit": r.returncode, "stdout": r.stdout, "stderr": r.stderr[-800:],
            "wrote": sorted(set(after) - set(before)),
            "changed": sorted(k for k in set(after) & set(before) if after[k] != before[k]),
            "vanished": sorted(set(before) - set(after)),
        })
    out = d / f"replay_{tag}.json"
    out.write_bytes(json.dumps(steps, ensure_ascii=False, indent=1).encode("utf-8"))
    return out


def cmd_replay(n: str) -> int:
    p = _replay_once(n, "a")
    steps = json.loads(p.read_bytes().decode("utf-8"))
    print(f"流し直した: {p}  （{len(steps)} コマンド）")
    for i, s in enumerate(steps, 1):
        print(f"  {i:2} exit={s['exit']:<2} {' '.join(s['argv'])[:72]}")
    return 0


# --- floor（揺れの床）------------------------------------------------------------------

def cmd_floor(n: str) -> int:
    """★★ **変更を入れずに 2 回**流し、その差を「床」とする。

    ★ なぜ要るか: 一段目は LLM で、同じ入力に同じ答えが返る保証が無い。
      この repo は 2026-08-24 に「同じブックに一字一句同じ依頼を 2 回投げると
      1 回目は通り 2 回目は落ちた」を実測している。2 体目の事故も揺れていた。
    ★ 床を知らずに差を読むと、**直したせいか、モデルの気分かを分けられない** ──
      2026-09-18 に「俺の検体が作った幻」を 2 件追いかけたのと同じ失敗を、
      次は 1 回ぶんぜんぶの規模でやることになる。
    """
    a = json.loads(_replay_once(n, "a").read_bytes().decode("utf-8"))
    b = json.loads(_replay_once(n, "b").read_bytes().decode("utf-8"))
    shaky = []
    for i, (x, y) in enumerate(zip(a, b), 1):
        why = []
        if x["exit"] != y["exit"]:
            why.append(f"exit {x['exit']}→{y['exit']}")
        if x["wrote"] != y["wrote"] or x["changed"] != y["changed"]:
            why.append("書いた物が違う")
        if why:
            shaky.append((i, " ".join(x["argv"])[:60], "／".join(why)))
    print(f"揺れの床: {len(shaky)} / {len(a)} コマンド")
    for i, cmd, why in shaky:
        print(f"  ★ {i:2} {cmd}  {why}")
    if not shaky:
        print("  ★ この回では揺れなかった ── ただし『床が 0』の証明ではない"
              "（測ったのはこの並びだけ）")
    (_run_dir(n) / "floor.json").write_bytes(
        json.dumps([{"step": i, "argv": c, "why": w} for i, c, w in shaky],
                   ensure_ascii=False, indent=1).encode("utf-8"))
    return 0


# --- sheet（採点表）---------------------------------------------------------------------

#: 機械で決まる条だけを見る。★ 意味の 4 条（中身が正しい／頼んでいないものを変えない／
#:   本当の理由を言う／文言が正確）は**人が読む** ── ここでは判定しない。
def cmd_sheet(n: str) -> int:
    d = _run_dir(n)
    p = d / "replay_a.json"
    if not p.is_file():
        print("× 先に replay を走らせてください")
        return 3
    steps = json.loads(p.read_bytes().decode("utf-8"))
    floor = {x["step"] for x in json.loads((d / "floor.json").read_bytes().decode("utf-8"))} \
        if (d / "floor.json").is_file() else set()
    sys.path.insert(0, str(REPO / "tests"))
    hard = []
    for i, s in enumerate(steps, 1):
        flags = []
        refused = s["exit"] != 0
        if refused and (s["wrote"] or s["changed"]):
            flags.append("断ったのに書いた（断り5）")
        if not refused and "？" in s["stdout"]:
            flags.append("断りの顔で exit 0（断り4）")
        if s["vanished"]:
            flags.append(f"消えた物: {s['vanished']}（到達3）")
        if refused and not any(k in s["stdout"] for k in ("--", "例:", "例：")):
            flags.append("逃げ道も通る書き方も示していない（断り3）")
        if flags:
            hard.append((i, " ".join(s["argv"])[:56], flags, i in floor))
    print(f"■ 機械で決まる 5 条だけの採点表（{len(steps)} コマンド）")
    print("★ 残り 4 条（中身が正しい／頼んでいないものを変えない／本当の理由を言う／"
          "文言が正確）は**人が読む** ── ここでは判定していない。")
    if not hard:
        print("  機械で拾える破れ: 0 件")
    for i, cmd, flags, on_floor in hard:
        mark = "（★ 揺れの床の上 ── 人が読むこと）" if on_floor else ""
        print(f"  {i:2} {cmd}{mark}")
        for f in flags:
            print(f"       ⚠ {f}")
    print("\n★ 合否はこの表では決まらない。9 条のうち 5 条を見ただけ"
          "（docs/FROZEN-20260917-盲検の合格線.md）。")
    return 0
