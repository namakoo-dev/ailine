"""本体と ailine_core が同じ名前を持つなら、**同じ物**を指す（2026-09-24・切り出しの主な合格線）。

★★ なぜ在るか: 関数を本文 1 文字も変えずに移しても、挙動が変わりうる道は 1 つだけ ──
  移した先で**名前が別の物に束縛される**こと（本体では A を指していた名前が、core では B を指す）。
  ゴールデンは通った行しか守らない（codegen の文の被覆は 77% だった）が、束縛は**全経路を静的に**見られる。
★ 見るのは**素の import の状態**（子プロセスで測る）。試験の切り離し（conftest）は走る前に
  ailine の属性を一時フォルダへ差し替えるので、試験の中で照合すると必ず外れる
  ── 2026-09-23 に差し替え名の番人がそれで変異に鳴らなかった。
★ 同じ名前の別物は、理由つきの在庫（KNOWN）にだけ許す。増えたら赤。在庫の古い項目も赤。
"""
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

#: ★ 同じ名前で別の物（2026-09-24 の棚卸し・7 件）。キーは「モジュール:名前」。
KNOWN = {
    "ailine_core.accounts_read:MAX_ROWS":
        "意味の違う定数（本体は表示の上限 1000・帳簿は読む上限 200000）。同じ名前にしているのは偶然。",
    "ailine_core.csv_quarantine:MAX_ROWS":
        "意味の違う定数（CSV の検疫の受理上限 50000）。本体の表示の上限とは別物と冒頭に書いてある。",
    "ailine_core.forms_collect:total_row":
        "本体の total_row は ailine_core.total_row モジュールで、こちらは合計行を返す関数。種類が違う。",
    "ailine_core.suggest:suggest_ops":
        "本体の suggest_ops は op の台帳を注いで core を呼ぶ薄い配線（引数の形が違う）。正当な別物。",
    "ailine_core.alias_store:_CJK_KANJI_RE":
        "★ 既知の欠陥の候補: 見た目は同じ漢字の範囲だが、パターンの文字列が一致しない。"
        "どちらが正しいかを確かめて 1 本に畳む（別 commit）。",
    "ailine_core.sum_identity:_is_number":
        "★ 既知の欠陥の候補: primitives.is_number（本体が _is_number として使う）と同じ名前で別の判断。"
        "primitives 自身が『同じ名前で違う判断』を警告している形。別 commit で名前か実装を揃える。",
}


def _measure() -> dict:
    probe = r'''
import importlib, json, pkgutil, sys, types
sys.path.insert(0, sys.argv[1])
import ailine, ailine_core
A = vars(ailine)
same, diff = 0, []
for m in pkgutil.walk_packages(ailine_core.__path__, "ailine_core."):
    mod = importlib.import_module(m.name)
    for n, v in vars(mod).items():
        if n.startswith("__") or n not in A or isinstance(v, types.ModuleType):
            continue
        if A[n] is v:
            same += 1
        else:
            diff.append(m.name + ":" + n)
print(json.dumps({"same": same, "diff": sorted(diff)}))
'''
    r = subprocess.run([sys.executable, "-c", probe, str(REPO / "src")], capture_output=True,
                       text=True, encoding="utf-8", timeout=120)
    assert r.returncode == 0, r.stderr[-1200:]
    return json.loads(r.stdout.strip().splitlines()[-1])


def test_a_shared_name_is_the_same_object():
    got = _measure()
    new = [k for k in got["diff"] if k not in KNOWN]
    assert not new, (
        "本体と ailine_core で同じ名前が別の物を指している ── 移した関数がその名前を使うと、"
        "本体と違う物を呼ぶ（本文が同じでも挙動が変わる）:\n  " + "\n  ".join(new)
        + "\n★ 移したなら本体は core の物を import し直す。意図した別物なら KNOWN に理由つきで")


def test_the_known_list_is_current_and_explained():
    got = _measure()
    stale = sorted(set(KNOWN) - set(got["diff"]))
    assert not stale, f"もう別物でない在庫が残っている（直したなら外す）: {stale}"
    for k, why in KNOWN.items():
        assert len(why) >= 30 and "。" in why, f"{k} の理由が薄い"


def test_the_measure_sees_shared_names():
    """★ 陽性対照: 共有している名前が数百ある（0 なら照合が空回りしている）。"""
    got = _measure()
    assert got["same"] >= 300, f"同じ物を指す共有名が {got['same']} しか無い ── 測れていない疑い"
