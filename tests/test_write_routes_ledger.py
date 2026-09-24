# 原本を書き換える**経路の台帳**（2026-09-02）── 全部がロックの関所を通ること。
#
# ★★ 発端: 自作 review が「`ailine redo` はロックの関所を一度も通っていない」を出した。
#   この repo は同型の事故（run は Excel ロックで止まるのに undo は素通り＝「復元の致命5」）
#   を既に踏み、番人の docstring に「**1 本で 4 経路を縛る**」と書いていた。
#   ★ それでも **5 本目（redo）を作って配線しなかった**。
#   ★ 「4 経路」という数は**人が数えて書いた**もので、増えても誰も気づかない形だった。
#
# ★★ この台帳を書きながら、俺の検出が浅いことも分かった（2026-09-02）:
#   最初「4 件が関所を通っていない」と出たが、それは**呼び元を 1 段しか辿らなかった**ため。
#   実際は `_cmd_run_body` が上流で通しており、**穴は無かった**。
#   ★ 数え方が浅いと「無い穴」を報告する ── 検出器も測定器として疑う。
#
# 契約:
#   ① 原本を書き換える口（低層）の顔ぶれが変わったら気づく
#   ② その口へ至る入口（コマンド層）が、全部ロックの関所を通る
#   ③ 関所の呼び出し箇所が減ったら退行（＝どこかの入口が素通りになった）

import inspect
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

from _product_source import product_files  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

#: ★ 2026-09-23: `inspect.getsource(ailine)` は**モジュール丸ごと＝本体 1 冊**だった。
#:   原本へ書く口が ailine_core へ移っても数えられるよう、製品全体を**ファイルごとに**行で持つ
#:   （持ち主の関数を上へ辿るとき、隣のファイルへはみ出さないように）。
FILE_LINES = [p.read_bytes().decode("utf-8").splitlines() for p in product_files()]


def _owners_of(lines: list) -> list:
    """行番号（0 起点）→ 持ち主の関数名。持ち主は**モジュール直下の関数か、クラス直下のメソッド**。

    ★ 2026-09-23: 初版は「行頭の def を上へ辿る」だった。GUI のメソッド（`Handler.do_POST`）の中の口を、
      上にある別の関数（`_default_dir`）の物と名指しした。字下げした def も数えるように緩めると、
      今度は内側の関数（`_finish_apply` の中の `_say`）を持ち主にした ── ★ 構文で決める。
    """
    import ast
    out = ["?"] * (len(lines) + 1)
    try:
        tree = ast.parse(chr(10).join(lines))
    except SyntaxError:
        return out
    fns = []
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            fns.append(n)
        elif isinstance(n, ast.ClassDef):
            fns += [m for m in n.body if isinstance(m, (ast.FunctionDef, ast.AsyncFunctionDef))]
    for f in fns:
        for k in range(f.lineno - 1, f.end_lineno):
            out[k] = f.name
    return out


OWNERS = [_owners_of(lines) for lines in FILE_LINES]


def _owner(lines: list, i: int) -> str:
    return OWNERS[next(k for k, ls in enumerate(FILE_LINES) if ls is lines)][i]


def _body(fn: str) -> str:
    """★ 2026-09-23: 名前で引く（本体の中の位置で切らない）。"""
    return inspect.getsource(getattr(ailine, fn))


# --- ① 原本へ実際にバイトを書く低層（ここが増えたら経路が増えたということ）--------------
WRITERS = {
    "_finish_apply": "run の適用結果を原本へ被せる（原子的置換）",
    "restore_backup": "undo/restore が世代から原本へ戻す",
    "redo_last_undo": "redo が退避から原本へ戻す",
    # ★★ 2026-09-23 に視野を GUI まで広げて初めて見えた口は、GUI の「原本に反映」（`do_POST`）が
    #   下書きを原本へ `shutil.copy2` で被せるだけ ── 関所もバックアップも通らない既知の欠陥だった。
    #   ★ 2026-09-24 に直した: 清書は `ailine adopt`（下の口）に畳み、GUI はそれを子プロセスで叩く。
    "cmd_adopt": "adopt が下書きの複写を原本へ被せる（原子的置換・控えあり）",
}

# --- ② 人の依頼が入る入口（ここが関所を通っていなければ素通り）--------------------------
ENTRIES = {
    "_cmd_run_body": "ailine run（1 冊）",
    "_cmd_undo_body": "ailine undo / restore",
    "cmd_redo": "ailine redo",
    "cmd_run_match": "ailine run（2 冊の突き合わせ）",
    "cmd_adopt": "ailine adopt（GUI の「原本に反映」もここを叩く）",
}


def _actual_writers() -> set:
    out = set()
    for lines in FILE_LINES:
        for i, ln in enumerate(lines):
            s = ln.strip()
            if s.startswith("#"):
                continue
            if "atomic_replace_inplace(" in ln and "def atomic_replace_inplace" not in ln:
                out.add(_owner(lines, i))
            if re.search(r"shutil\.copy2\([^)]*,\s*book\)", ln):
                out.add(_owner(lines, i))
    return out


def _actual_gate_callers() -> set:
    return {_owner(lines, i) for lines in FILE_LINES for i, ln in enumerate(lines)
             if "refuse_if_locked(" in ln and "def refuse_if_locked" not in ln}


def test_the_set_of_writers_has_not_changed():
    """① 原本へ書く低層の顔ぶれ ── 増えたら新しい経路ができたということ。"""
    now = _actual_writers()
    added = sorted(now - set(WRITERS))
    gone = sorted(set(WRITERS) - now)
    assert not added, (
        f"原本へ書く新しい口ができた: {added} ── その口へ至る入口が"
        "ロックの関所を通るか確かめ、WRITERS に足すこと")
    assert not gone, f"原本へ書く口が消えた: {gone}（台帳を更新すること）"


def test_every_entry_goes_through_the_lock_gate():
    """② 入口が全部、関所を通ること。

    ★ ここが今日 review に突かれた所（redo が 5 本目の入口として素通りしていた）。
    """
    lack = [fn for fn in ENTRIES if "refuse_if_locked" not in _body(fn)]
    assert not lack, f"ロックの関所を通っていない入口: {lack} ── {[ENTRIES[f] for f in lack]}"


def test_the_gate_is_not_called_from_fewer_places():
    """③ 関所の呼び出しが減ったら退行（どこかの入口が素通りになった）。"""
    now = _actual_gate_callers()
    gone = sorted(set(ENTRIES) - now)
    assert not gone, f"関所を呼ばなくなった入口: {gone}"


def test_the_ledger_counts_are_stated():
    """★ 「4 経路」のような数を**人が書いて古くする**のを止める。

    ★ docstring に数を書くなら、その数を機械が確かめる側が要る
      （書いた数と実体がずれるのが、この repo が何度も踏んだ形）。
    """
    now = _actual_gate_callers() & set(ENTRIES)
    assert len(now) == len(ENTRIES), (
        f"入口 {len(ENTRIES)} 件のうち関所を通るのは {len(now)} 件")
    body = _body("refuse_if_locked")
    # ★★ 2026-09-22（番人の生存試験で見つけた）: ここは `if "経路を縛る" in body:` で
    #   **条件つき**だった。字義どおりの契約（数が書いてあるなら検算する）には合っている。
    #   ★ 危ないのは**言い換え**の方 ── 文言を「4 つの経路を守る」などに直すと、
    #     数は残ったまま検査だけが静かに消える。今日、画面の文言を 2 度言い換えた。
    #   ★ だから宣言を**必須**にする。消す・言い換えるときは、ここで大声で止まる。
    assert "経路を縛る" in body, (
        "refuse_if_locked の docstring から『N 経路を縛る』の宣言が消えています ── "
        "言い換えるなら、この番人の読み方も一緒に直してください（数だけ残すと古くなる）")
    m = re.search(r"(\d+)\s*経路を縛る", body)
    assert m and int(m.group(1)) == len(ENTRIES), (
        f"docstring の経路数（{m.group(1) if m else '?'}）と実体（{len(ENTRIES)}）が違う")
