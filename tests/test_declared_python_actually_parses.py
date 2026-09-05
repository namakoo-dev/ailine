"""pyproject が宣言した Python でも、ソースが**構文として通る**こと（2026-09-05）。

★★ 出所（この日の実装中に自分で踏んだ）: 属性の登録を書いていて、こう書いた ──

    return False, f"属性『{getattr(cand, "kind", None)}』は登録できません"

  f 文字列の中に**同じ引用符**を入れる書き方は **3.12 以降しか通らない**（PEP 701）。
  pyproject は `requires-python = ">=3.10"` と宣言している。手元も CI も 3.12 なので
  **誰にも鳴らないまま出荷される**ところだった。

★ 「在っても鳴らない」の形そのもの ── 番人が無いのではなく、宣言（3.10 以上）と
  実際に走らせている環境（3.12 だけ）が食い違っていて、その事故の形では誰も見ていない。

★★ この番人自身が一度**空振りした**（同じ日・この場で）:
  最初は `ast.parse(..., feature_version=(3, 10))` だけで守ったつもりでいた。
  自分に向けた変異試験（下の test_the_guard_catches_...）を書いたら**落ちた** ──
  feature_version は PEP 701 の入れ子引用符を**捕まえない**。
  ★ 「効かなかった」ではなく「測れていなかった」。トークン列を直接見る形に替えた。

★ 射程を正直に言う: ここが見るのは**構文だけ**。新しい標準ライブラリの関数を呼んだ、
  といった**実行時**の非互換は捕まえられない。それを見たければ CI に 3.10 を足すしかない
  （★ 発火条件つきの保留: 3.10 を買い手が実際に使うと分かったら、CI に足す。
    いま足さないのは「赤くなった時に直す余力を今日は持っていない」だけで、
    技術的な行き止まりではない）。
"""
import ast
import io
import re
import sys
import tokenize
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

#: f 文字列の中に同じ引用符を入れられるようになった版（PEP 701）
PEP701 = (3, 12)


def _declared_minimum() -> tuple:
    text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    m = re.search(r'requires-python\s*=\s*"[^"]*?(\d+)\.(\d+)', text)
    assert m, "pyproject.toml に requires-python が読めない"
    return int(m.group(1)), int(m.group(2))


def quotes_nested_in_fstrings(src: str) -> list:
    """f 文字列の中で**同じ引用符**を使っている箇所の行番号（PEP 701 以降でしか通らない）。

    ★ ast の feature_version では捕まらない（実測 ── この番人が一度空振りした）。
      トークン列で f 文字列の開始/終了を追い、その内側の文字列トークンの引用符を見る。
    ★ 走行中の Python が 3.12 未満なら、そもそもこの構文はここで SyntaxError になるので
      検出は不要（空を返す）。
    """
    start = getattr(tokenize, "FSTRING_START", None)
    end = getattr(tokenize, "FSTRING_END", None)
    if start is None or end is None:
        return []
    hits, stack = [], []
    try:
        for tok in tokenize.generate_tokens(io.StringIO(src).readline):
            if tok.type == start:
                stack.append(tok.string[-1])          # f" なら "、f' なら '
            elif tok.type == end:
                if stack:
                    stack.pop()
            elif tok.type == tokenize.STRING and stack:
                if tok.string.lstrip("rbfRBF")[:1] == stack[-1]:
                    hits.append(tok.start[0])
    except (tokenize.TokenError, IndentationError, SyntaxError):
        return []
    return hits


def test_pyproject_declares_a_minimum():
    major, minor = _declared_minimum()
    assert major == 3 and minor >= 8, (major, minor)


@pytest.mark.parametrize("path", sorted(SRC.rglob("*.py")), ids=lambda p: p.name)
def test_every_source_file_works_on_the_declared_minimum(path):
    """★ 宣言した下限で ①構文が通る ②PEP 701 の書き方を使っていない。"""
    minimum = _declared_minimum()
    src = path.read_text(encoding="utf-8")
    try:
        ast.parse(src, filename=str(path), feature_version=minimum)
    except SyntaxError as e:
        pytest.fail(f"{path.relative_to(ROOT)}:{e.lineno} は Python "
                    f"{minimum[0]}.{minimum[1]} では通らない: {e.msg}")
    if minimum < PEP701:
        bad = quotes_nested_in_fstrings(src)
        assert not bad, (f"{path.relative_to(ROOT)}:{bad} は f 文字列の中に同じ引用符を"
                         f"使っている（{PEP701[0]}.{PEP701[1]} 以降でしか通らないのに、"
                         f"pyproject は {minimum[0]}.{minimum[1]} 以上と宣言している）")


def test_the_guard_catches_the_thing_that_bit_us():
    """★ 番人自身への変異試験 ── これが最初の版を落とした（feature_version では無理だった）。"""
    bitten = 'x = f"{getattr(o, "k", None)}"' + chr(10)
    fine = 'x = f"{getattr(o, ' + chr(39) + 'k' + chr(39) + ', None)}"' + chr(10)
    assert quotes_nested_in_fstrings(bitten) == [1]
    assert quotes_nested_in_fstrings(fine) == []


def test_the_guard_does_not_cry_wolf_on_ordinary_code():
    """★ 誤爆側を対で縛る（在るのに『無い』と言う方より、無いのに騒ぐ方が邪魔）。"""
    for ok in ('x = f"{a}{b}"', "y = f'{a}' + " + chr(34) + "b" + chr(34),
               'z = "plain" + f"{n}"', 'w = f"{d[' + chr(39) + 'k' + chr(39) + ']}"'):
        assert quotes_nested_in_fstrings(ok + chr(10)) == [], ok


def test_ci_runs_the_version_we_declare():
    """★ この穴が**二度と静かに開かない**ための番人（2026-09-05・Namakoo 決裁）。

      宣言は「3.10 以上」、CI は 3.12 だけ ── その隙間に、3.12 でしか通らない書き方が
      **23 箇所**溜まっていた。3.10/3.11 の買い手は import で落ちる。誰も見ていなかった。
    ★ 宣言を実測（3.12）に寄せた。以後は「宣言した下限を CI が実際に走らせている」ことを
      機械が突き合わせる ── 下限を下げるなら CI にその版を足すのが先。
    """
    declared = _declared_minimum()
    wf = ROOT / ".github" / "workflows" / "tests.yml"
    if not wf.is_file():
        pytest.skip("tests.yml が無い")
    versions = {tuple(int(x) for x in v.split("."))
                for v in re.findall(r'python-version:\s*"?(\d+\.\d+)"?', wf.read_text(encoding="utf-8"))}
    assert versions, "CI が python-version を指定していない"
    assert declared in versions, (
        f"宣言した下限 {declared[0]}.{declared[1]} を CI が走らせていない（CI: {sorted(versions)}）"
        "── 下限を下げるなら、その版を CI に足してから")


def test_the_running_interpreter_is_not_the_only_thing_checked():
    """★ 「手元で動いた」を根拠にしていないこと（走行中の版と宣言の下限は別物）。"""
    assert sys.version_info[:2] >= _declared_minimum()
