"""「こう打てば進める」と言った所は、全部台帳に載っていること（2026-09-23）。

★★ 事故（盲検 6 体目 ⑥・致命）: 道具が「`--column 借方勘定科目=科目` を付けて
  もう一度実行してください」と案内し、**その通りに打つと落ちた**。
  ★ ⑦ も同じ形（「『取引先をキーに』と書け」と言われて、既に書いてある）。

★★ 「断りが示した道を実際に歩く」道具は**既に在った**（tests/walk_refusals_core.py）。
  判定の芯に `path_fails`（歩いたが着かなかった＝**通らない道を示した**）まで持っている。
  ★ なのに ⑥ が素通りした ── **台帳（refusal_register.json）に accounts が 1 件も
    載っていなかった**から。

★★ 根は分母の作り方だった。台帳の `_meta.denominator` はこう書いてある:

    「src/ の print/say に渡る文字列で**『？』から始まるもの**」

  ★ `accounts` の断りは `×` で始まり、しかも**道を示す文は別の関数で組み立てて
    後から連結**される。だから『？』でも『×』でも拾えない ──
    **導線の半分が、構造的に分母の外に在った。**

★ ここが数えるのは「**打てるものを名指ししている文字列**」──
  `` `ailine <サブコマンド>` `` か `` `--フラグ` `` をバッククォートの中に含むもの。
  実測（2026-09-23）で **40 件**。そのすべてが既存の台帳の外だった。

★ この番人は「歩けるか」までは見ない（それは walk_refusals の仕事）。
  ★ まず**在ることを数える** ── 分母が無ければ、歩く先も決まらない。
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _product_source import product_strings  # noqa: E402

REGISTER = Path(__file__).resolve().parent / "typable_hints_register.json"

#: ★ 「打てるもの」= この道具のサブコマンドか、長いフラグ。バッククォートの中だけを見る
#:   （散文の中の `--` はコマンドでないことがある）。
TYPABLE = re.compile(r"`(ailine [a-z-]+|--[a-z-]+)")


def typable_hints() -> list:
    """(ファイル, 行, 打てるもの, 文字列) を全部。★ 分母は機械が出す。"""
    out = []
    for f, ln, v in product_strings():
        m = TYPABLE.search(v)
        if m:
            out.append((Path(f).name, ln, m.group(1), v))
    return out


def _register() -> dict:
    return json.loads(REGISTER.read_text(encoding="utf-8"))


def test_the_register_exists_and_says_how_the_denominator_is_made():
    """★ 分母の作り方が台帳に書いてあること（手書きの名簿にしない）。"""
    d = _register()
    assert "denominator" in d["_meta"], "分母の作り方が書かれていない"
    assert "打てる" in d["_meta"]["denominator"], d["_meta"]["denominator"]


def test_every_typable_hint_is_in_the_register():
    """★★ 本体 ── 「こう打てば進める」と言った所は、全部台帳に在ること。

    ★ 足す時は `walked`（その道を実際に歩いて確かめたか）を書く。
      書けないなら `未調査` と書く ── **見ていないものを「見た」にしない**。
    """
    known = {f"{e['file']}:{e['line']}" for e in _register()["hints"]}
    now = {f"{name}:{ln}" for name, ln, _t, _v in typable_hints()}
    missing = sorted(now - known)
    assert not missing, (
        f"台帳に無い導線が {len(missing)} 件あります: {missing[:6]}" + chr(10)
        + "  ★ 『こう打てば進める』と書いたなら、その道が通ることを誰かが確かめる"
        " 必要があります。tests/typable_hints_register.json に足してください")


def test_the_register_does_not_keep_stale_rows():
    """★ 実体から消えた導線を台帳に残さない（古い不安を配らない）。"""
    now = {f"{name}:{ln}" for name, ln, _t, _v in typable_hints()}
    stale = sorted({f"{e['file']}:{e['line']}" for e in _register()["hints"]} - now)
    assert not stale, f"実体に無い行が台帳に残っています: {stale[:6]}"


def test_every_row_says_whether_anyone_walked_it():
    """★ 「歩いたか」を必ず書く ── 空欄を「大丈夫」と読ませない。"""
    ok = {"walked", "path_fails", "by_design", "未調査"}
    bad = [e for e in _register()["hints"] if e.get("walked") not in ok]
    assert not bad, (
        f"walked が {sorted(ok)} のどれでもない行: "
        f"{[(e['file'], e['line'], e.get('walked')) for e in bad][:6]}")


def test_the_one_that_bit_us_is_marked():
    """★ 実際に事故った導線（accounts の --column）が、そう記録されていること。

    ★ 記録が「直したから消す」になっていないか ── 出所は残す。
    """
    rows = [e for e in _register()["hints"] if e["file"] == "accounts_core.py"]
    assert rows, "accounts_core.py の導線が 1 件も載っていない"
    assert any("6 体目" in (e.get("note") or "") for e in rows), (
        "★ 盲検 6 体目で事故った導線だと分かる記録が無い")


def _declared():
    """argparse の宣言から、サブコマンドと長いフラグを取る（手書きしない）。"""
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    import ailine
    subs, flags = set(), set()
    for a in ailine.build_parser()._actions:
        if hasattr(a, "choices") and a.choices and hasattr(a.choices, "keys"):
            for name, sp in a.choices.items():
                subs.add(name)
                flags |= {o for act in sp._actions
                          for o in act.option_strings if o.startswith("--")}
    return subs, flags


def test_every_advertised_name_actually_exists():
    """★★ 案内した名前が**宣言に在る**こと（打てないものを勧めない）。

    ★ これは**存在の確認であって、歩けることの確認ではない** ──
      盲検 6 体目 ⑥ の `--column` は**実在したうえで落ちた**。
      歩行は walked 欄（と walk_refusals）の仕事。ここは安い側だけを機械で縛る。
    ★ 分母は argparse の宣言から引く（名簿を手で持たない）。
    """
    subs, flags = _declared()
    bad = []
    for name, ln, typ, _v in typable_hints():
        if typ.startswith("ailine "):
            if typ.split()[1] not in subs:
                bad.append((name, ln, typ, "そんなサブコマンドは無い"))
        elif typ not in flags:
            bad.append((name, ln, typ, "そんなフラグはどのサブコマンドにも無い"))
    assert not bad, f"実在しないものを案内しています: {bad[:6]}"


def test_existing_is_not_the_same_as_walkable():
    """★ 「在る」と「歩ける」を混ぜない ── 台帳にまだ歩いていない行が在ることを認める。

    ★ ここが緑でなくなったら、38 件を歩き終えたということ（その時はこの試験を消す）。
      ★ 0 件でも通る試験にしない ── 数が減ったら**気づく**側に倒す。
    """
    todo = [e for e in _register()["hints"] if e.get("walked") == "未調査"]
    assert todo, "★ 未調査が 0 件 ── 歩き終えたならこの試験は消してよい"
    assert len(todo) <= 38, (
        f"未調査が増えています（{len(todo)} 件）── 新しい導線を足したなら歩いてから")
