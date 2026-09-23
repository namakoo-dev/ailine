"""ailine_core は、試験が `ailine` に差し替えている名前を**自分の中で使わない**（2026-09-23・切り出しの番人）。

★★ なぜ在るか（設計レビューで第 1 版の地図が崩れた件）:
  試験は偽物を `ailine` の属性に入れる（`monkeypatch.setattr(ailine, "X", ...)`）。本体の関数は
  呼ぶ時に `ailine` の大域から X を引くので届く。ところが X の実体を ailine_core に移し、
  core の中の関数が X を使うと、**core は自分の束縛を引く** ── 差し替えは届かず、
  試験は本物（ollama・~/.ailine の下の保存先）に触ったまま**緑**になりうる。
  例: `load_vocab` は `VOCAB_FILE`（~/.ailine の下）を直に読む。これを core へ移すと、
  試験の切り離しが黙って外れ、本物のホームに書く（今日、別の経路で実害が出た形）。
★ 名簿は手で書かない: 差し替え名は試験の setattr から、ホームの下の保存先は本体から導く
  （tests/split_progress_core.py と同じ数え方 ── 測定器を 1 つにする）。
★ 見るのは「`ailine.X` が core のある束縛と**同じ実体**（引き継ぎ）」の時だけ。同じ名前の別物
  （`MAX_ROWS` は本体 1000・accounts_read 200000・csv_quarantine 50000 ── 意味が違う定数）は当たらない。
"""
import ast
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))
import split_progress_core as spc  # noqa: E402

CORE = HERE.parent / "src" / "ailine_core"

#: ★ 理由つきで許すもの（キー: 名前）。理由は 30 字以上・句点あり。
ALLOWED = {
    "_used_extent":
        "試験の差し替えは本体の _sheet_with_nothing_in_it の故障注入（読んだ後に落ちる回）で、"
        "core の事後条件（derive）の呼び出しを止める意図はない。本体の関数が core へ移る日は"
        "この免除を外して試験の的を core に移すこと。",
}


def core_uses(core_dir: Path = CORE) -> dict:
    """{名前: [(ファイル, 行)]} ── core の中で Load されている名前（トップレベルの定義も含む）。"""
    out = {}
    for p in sorted(core_dir.rglob("*.py")):
        tree = ast.parse(p.read_bytes().decode("utf-8"))
        for x in ast.walk(tree):
            if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load):
                out.setdefault(x.id, []).append((p.relative_to(core_dir).as_posix(), x.lineno))
    return out


def inherited_from_core(names: set) -> set:
    """本体が `from ailine_core... import X` で引き継いでいる名前（★ 静的に読む）。

    ★★ 初版は実行時に `ailine.X is core.X` で照合していて、**変異試験で鳴らなかった**（2026-09-23）:
      試験の切り離し（conftest）が走る前に `ailine.VOCAB_FILE` を一時フォルダの Path に差し替えるので、
      照合は必ず外れる ── 番人を守る仕組みが番人の目を塞いでいた。★ 本体の import 文を読む。
    """
    tree = ast.parse(spc.MAIN.read_bytes().decode("utf-8"))
    out = set()
    for x in ast.walk(tree):
        if (isinstance(x, ast.ImportFrom) and x.level == 0
                and (x.module or "").split(".")[0] == "ailine_core"):
            out |= {a.asname or a.name for a in x.names}
    return out & names


def hazards(patched: set, inherited: set, uses: dict) -> dict:
    """★ 純関数（変異試験で直接叩く）: 差し替えられ・core から引き継がれ・core の中で使われる名前。"""
    return {nm: uses[nm] for nm in sorted(patched & inherited) if nm in uses}


def _held_names() -> set:
    return spc.patched_names() | spc.home_names()


def test_core_does_not_use_a_name_the_tests_patch():
    found = hazards(_held_names(), inherited_from_core(_held_names()), core_uses())
    bad = {k: v for k, v in found.items() if k not in ALLOWED}
    assert not bad, (
        "試験が ailine に差し替える名前を ailine_core の中で使っている ── 差し替えが届かず、"
        "試験が本物に触ったまま緑になりうる:\n" + "\n".join(
            f"  {k}: {v[:4]}" for k, v in bad.items())
        + "\n★ 直し方: 作用（保存先・翻訳・LibreOffice）は core に置かず、呼び出し時に引数で受ける"
          "（前例 DslStepDeps）。移さないなら本体に残す。")


def test_the_allowances_are_still_needed_and_explained():
    found = hazards(_held_names(), inherited_from_core(_held_names()), core_uses())
    stale = sorted(set(ALLOWED) - set(found))
    assert not stale, f"もう当たらない免除が残っている（古い不安を配らない）: {stale}"
    for k, why in ALLOWED.items():
        assert len(why) >= 30 and "。" in why, f"免除 {k} の理由が薄い"


def test_the_measure_sees_the_patches():
    """★ 陽性対照: 差し替え名が 0 なら上の検査は空回りする（setattr を拾えていない）。"""
    held = _held_names()
    assert len(held) >= 20, f"差し替え名が {len(held)} 個しか拾えていない"
    assert "VOCAB_FILE" in held, "ホームの下の保存先（VOCAB_FILE）を拾えていない"
    assert inherited_from_core(held), "core から引き継いだ名前が 0 ── 実体の照合が壊れている疑い"


def test_the_detector_catches_the_load_vocab_shape():
    """★ 変異: `load_vocab` を core へ移した形（VOCAB_FILE を core が持ち、中で読む）を食わせると鳴る。"""
    uses = {"VOCAB_FILE": [("vocab.py", 12)], "re": [("x.py", 1)]}
    assert hazards({"VOCAB_FILE", "translate_task"}, {"VOCAB_FILE"}, uses) == {
        "VOCAB_FILE": [("vocab.py", 12)]}
    # 引き継ぎでない同名（別物の定数）は鳴らない
    assert hazards({"MAX_ROWS"}, set(), {"MAX_ROWS": [("accounts_read.py", 62)]}) == {}
