# #15 run.lock の居座り ── 修正より先に凍結した赤い検体。
#
# 実測した事故: 前の ailine が終わった後、その PID を**無関係な別プロセスが取り直す**と
# 「まだ生きている」と誤判定し、ロックが居座って直後の run が exit 6 になる。
# ★ 根: 判定に要る三項（誰のロックか／今その PID は誰か／同じか）のうち、
#   **誰のロックか**を持っていなかった（tasklist の出力に PID の数字が含まれるか、
#   という部分文字列判定だった）。
#
# 契約:
#   ① 取得したロックには「誰か」（実行ファイル名）が焼かれている
#   ② PID が同じでも**別のプロセス**なら stale として奪える
#   ③ 自分自身のロックは当然 stale でない（誤爆で二重実行を許さない）
#   ④ 古い形式（image が無いロック）でも壊れない（後方互換・安全側）

import json
import os
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def test_lock_records_who_holds_it(tmp_path):
    """① 誰のロックかを焼く（第三項）。"""
    p = tmp_path / "run.lock"
    acquired, msg = ailine.acquire_run_lock(p)
    assert acquired, msg
    info = json.loads(p.read_text(encoding="utf-8"))
    assert info["pid"] == os.getpid()
    assert info.get("image"), "誰のロックかが焼かれていない"
    ailine.release_run_lock(p)


def test_same_pid_different_process_is_stale():
    """② PID の使い回し ── 名前が違えば「もう居ない」と判断して奪える。"""
    info = {"pid": os.getpid(), "ts": "2026-08-24T00:00:00+00:00",
            "image": "definitely_not_us.exe"}
    assert ailine._lock_is_stale(info), \
        "PID の使い回しで居座る（直後の run が exit 6 になる）"


def test_our_own_live_lock_is_not_stale():
    """③ 誤爆防止: 本物の生きたロックを奪わない（二重実行の防止が本業）。

    ★ 初版は「pid+1」を装って測ろうとしたが、その pid が実在せず skip になった。
      skip は『測れなかった』であって『安全』ではない ── **本物の同種プロセスを
      立てて**測る（陽性対照）。
    """
    import subprocess
    child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    try:
        assert ailine._pid_alive(child.pid, expect_image=ailine.own_image_name()),             "生きている同種プロセスを『居ない』と言った（測定器の故障）"
        info = {"pid": child.pid,
                "ts": ailine.datetime.now(ailine.timezone.utc).isoformat(),
                "image": ailine.own_image_name()}
        assert not ailine._lock_is_stale(info), "生きた同種プロセスのロックを奪おうとした"
        # 感度: 名前が違えば同じ pid でも stale（この判定が効いていることの確認）
        other = dict(info, image="definitely_not_us.exe")
        assert ailine._lock_is_stale(other), "名前の照合が効いていない（感度ゼロ）"
    finally:
        child.kill()
        child.wait(timeout=10)


def test_old_format_lock_still_works(tmp_path):
    """④ image を持たない古いロックでも例外にならない（後方互換・安全側）。"""
    info = {"pid": 999999, "ts": "2026-08-24T00:00:00+00:00"}
    assert ailine._lock_is_stale(info) is True     # 実在しない pid → 奪える
    info2 = {"pid": os.getpid(), "ts": "2026-08-24T00:00:00+00:00"}
    assert ailine._lock_is_stale(info2) is True    # 自分自身 → 従来どおり stale


def test_dead_pid_is_not_reported_alive():
    """恒真殺し: まず在り得ない PID を「生きている」と言わない。"""
    assert ailine._pid_alive(999999, expect_image=ailine.own_image_name()) is False
