"""bench の道具は本物の ~/.ailine に書かない ── 入口は ここ 1 つ（2026-09-24）。

★★ なぜ在るか（実害）: 2026-09-24 に効果の行列（basic_ops_matrix.py）を pytest の外で直に流したら、
  本物の ~/.ailine/history.jsonl に 246 行を書いた。道具は `python -m ailine run …` を子プロセスで呼び、
  環境変数をそのまま渡していた ── 試験（conftest）は切り離すが、bench の道具は誰も切り離していなかった
  （製品を動かす bench の道具 19 本すべて）。
★ 使い方: 道具の**先頭**（docstring と `from __future__` の直後・ailine を import する前）で

      import _bench_home; _bench_home.isolate()

  AILINE_HOME が一時フォルダになり、子プロセス（`-m ailine`）にもそのまま継がれる。同じプロセスで
  ailine を import する道具も、import の時にこの AILINE_HOME から保存先を決める。
★ 既に AILINE_HOME がある時（pytest の中・人が指定した時）は何もしない。
★ 本物の履歴を**読む**のが目的の道具（residue_gate_on_history.py）は、自分でパスを書いて読むので影響を受けない。
★ 番人: tests/test_bench_tools_do_not_write_the_real_home.py（製品を動かす道具の名簿はコードから導く）。
"""
from __future__ import annotations

import atexit
import os
import shutil
import sys
import tempfile


def isolate() -> str:
    """AILINE_HOME を一時フォルダにする（終わったら消す）。戻り値はそのフォルダ。"""
    if os.environ.get("AILINE_HOME"):
        return os.environ["AILINE_HOME"]
    if "ailine" in sys.modules:
        raise RuntimeError("_bench_home.isolate() は ailine を import する前に呼ぶこと"
                           "（import の時に保存先が決まるので、後からでは本物のホームを指したまま）")
    d = tempfile.mkdtemp(prefix="ailine_bench_home_")
    os.environ["AILINE_HOME"] = d
    atexit.register(shutil.rmtree, d, True)
    return d
