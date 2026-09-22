"""番人の生存試験 ── 守ると称する契約を破ったとき、その番人は本当に赤くなるか。

★★ なぜ在るか（2026-09-22 の実測）: 画面の語を 2 度言い換えた日に、**字面で書かれた
  番人を 3 本**見つけた。うち 1 本（盲検の解析器）は拾えなくなっても**黙った** ──
  assert ではなく正規表現だったので、例外にならず「着いた先ゼロ」を静かに返す。
  ★ 残る 2 本は大声で落ちた。つまり危ないのは「字面であること」そのものではなく:
      ① 解析器・抽出器（拾えなくても黙る）
      ② 否定形の検査（無いことを見ているので、消えると通る）
      ③ 字面だけ直して契約を確かめない人の手当て
  だから数えるべきは字面の数ではなく、**変異を入れたとき赤くなるか**そのもの。

★ 何を変異させるか: **人に見せる文言の字面**だけ。コードの形（`def foo(` など）は
  消すと構文や import が壊れ、「番人が捕まえた」のか「壊れて落ちた」のか区別できない。
  文言なら中身を変えても構文は保たれるので、**赤＝番人が観測している**と言い切れる。

★ 安全の作法:
  - 作業木が汚れていたら走らせない（戻し損ねても git で取り返せる状態でだけ動く）
  - 変異 → 試験 → 復元 の復元は**バイトで検算**する
  - 1 つの needle ごとに即復元する（途中で落ちても汚れを残さない）
"""
import hashlib
import io
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

HELPERS = {"count_in_product", "count_in_code", "window_around"}
MANGLE = "ZZ番人の生存試験ZZ"


def _has_japanese(s: str) -> bool:
    return any("぀" <= c <= "ヿ" or "一" <= c <= "鿿" for c in s)


def _looks_like_code(s: str) -> bool:
    t = s.strip()
    return (t.startswith("def ") or t.startswith("if ") or t.startswith("_")
            or t.endswith("(") or "=" in t.split("(")[0])


def wording_needles() -> list:
    """試験が**人に見せる文言**を字面で持っている所を (試験ファイル, 行, 文言) で返す。"""
    import ast
    out = []
    for p in sorted((REPO / "tests").glob("*.py")):
        try:
            tree = ast.parse(io.open(p, encoding="utf-8").read())
        except SyntaxError:
            continue
        for n in ast.walk(tree):
            if not (isinstance(n, ast.Call) and isinstance(n.func, ast.Name)
                    and n.func.id in HELPERS and n.args):
                continue
            a = n.args[0]
            if not (isinstance(a, ast.Constant) and isinstance(a.value, str)):
                continue
            v = a.value
            if _has_japanese(v) and not _looks_like_code(v):
                out.append((p, n.lineno, v))
    return out


def _product_paths() -> list:
    from _product_source import product_files
    return list(product_files())


def _sha(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()[:16]


def mutate(needle: str) -> dict:
    """文言を書き換える。戻すための {パス: 元のバイト} を返す（当たらなければ空）。"""
    saved = {}
    for p in _product_paths():
        b = io.open(p, "rb").read()
        s = b.decode("utf-8")
        if needle not in s:
            continue
        saved[p] = b
        io.open(p, "wb").write(s.replace(needle, MANGLE).encode("utf-8"))
    return saved


def restore(saved: dict) -> None:
    for p, b in saved.items():
        io.open(p, "wb").write(b)
        got = io.open(p, "rb").read()
        assert _sha(got) == _sha(b), f"復元できていない: {p}"


def run_one(test_file: Path, timeout=900) -> bool:
    """その試験ファイルが緑か。"""
    r = subprocess.run(
        [sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", "--no-header",
         "-x", "-m", "not local", str(test_file)],
        cwd=str(REPO), capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=timeout)
    return r.returncode == 0


def tree_is_clean() -> bool:
    r = subprocess.run(["git", "status", "--porcelain"], cwd=str(REPO),
                       capture_output=True, text=True)
    return r.returncode == 0 and not r.stdout.strip()


def main():
    if not tree_is_clean():
        print("✗ 作業木が汚れています。戻し損ねたときに取り返せないので走らせません。")
        return 2
    needles = wording_needles()
    print(f"人に見せる文言を字面で持つ番人: {len(needles)} 件")
    print()
    alive, dead, skipped = [], [], []
    for i, (tf, ln, needle) in enumerate(needles, 1):
        head = f"[{i}/{len(needles)}] {tf.name}:{ln}"
        print(f"{head}  {needle[:40]!r}", flush=True)
        if not run_one(tf):
            skipped.append((tf, ln, needle, "変異前から赤"))
            print("    … 変異前から赤（測れない）", flush=True)
            continue
        saved = mutate(needle)
        if not saved:
            skipped.append((tf, ln, needle, "製品に見つからない"))
            print("    … 製品にこの文言が無い（★ 既に空振りの可能性）", flush=True)
            continue
        try:
            green = run_one(tf)
        finally:
            restore(saved)
        if green:
            dead.append((tf, ln, needle))
            print("    ★ 緑のまま ── この番人はこの文言を見ていない", flush=True)
        else:
            alive.append((tf, ln, needle))
            print("    ✓ 赤になった", flush=True)

    n = len(alive) + len(dead)
    print()
    print(f"生存 {len(alive)} / 測れた {n}"
          + (f" = {len(alive)/n:.0%}" if n else ""))
    if dead:
        print("★ 変異しても緑だった番人:")
        for tf, ln, v in dead:
            print(f"   {tf.name}:{ln}  {v[:50]!r}")
    if skipped:
        print("測れなかったもの:")
        for tf, ln, v, why in skipped:
            print(f"   {tf.name}:{ln}  {why}  {v[:40]!r}")
    assert tree_is_clean(), "★ 作業木が汚れたまま終わった ── git status を見てください"
    print("作業木は元どおり（検算済み）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
