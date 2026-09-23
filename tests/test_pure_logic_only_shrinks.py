"""本体（src/ailine/__init__.py）に残っている**純ロジック**は、減る向きにしか動かない（2026-09-23）。

★★ なぜ在るか: 分割の目的は「後の保守と切り分け・当該部分だけを持ち出せること」
  （test_line_budget.py 冒頭の Namakoo の言:「太ること自体は問題ない。筋肉なら良い。贅肉で太るのは困る」）。
  ★ 行数の上限は 08-16 に意図して捨てた ── 配線（筋肉）まで縛って、空行を削る整形を強いたから。
  ここが縛るのは**配線でない方**だけ: I/O を持たず、試験が差し替える名前も読まない関数
  （＝ ailine_core へ持ち出せる部品）が、本体の中で**増えないこと**。
★ 数え方は tests/split_progress_core.py の 1 か所（scripts/split_progress.py と同じ測定器）。
★ 記録（tests/ailine_pure_logic_ceiling.txt）は `scripts/refresh_records.py --write` が
  **下げる向きにだけ**書き直す。上げるのは人が手で ── 理由を commit に書く（黙って上がらない）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import split_progress_core as spc  # noqa: E402

CEILING_FILE = Path(__file__).resolve().parent / "ailine_pure_logic_ceiling.txt"


def _ceiling() -> int:
    return int(CEILING_FILE.read_bytes().decode("utf-8").split()[0])


def test_pure_logic_in_the_main_module_does_not_grow():
    lines, n = spc.pure_logic(spc.survey())
    ceiling = _ceiling()
    assert lines <= ceiling, (
        f"本体の純ロジックが増えた（{lines} 行 / 記録 {ceiling} 行・{n} 関数）── "
        "I/O も差し替え名も持たない関数は ailine_core に置く。どうしても本体に要るなら "
        "記録を手で上げ、理由を commit に書く。"
        "（いま純と数えた関数は `python scripts/split_progress.py --pure` で見られる）")


def test_the_measure_is_alive():
    """★ 陽性対照 ── 測定器が壊れて 0 を返していないこと。

    既知の純な関数（コード生成の 1 本）が純と数えられ、既知の不純な関数（ホームの下を読む）が
    不純と数えられること。★ 片方だけだと「全部純」「全部不純」に壊れても通る。
    """
    funcs = spc.survey()["funcs"]
    pure = [k for k, v in funcs.items() if v["pure"]]
    assert pure, "純と数えた関数が 0 ── 測定器が壊れている疑い"
    assert any(not v["pure"] for v in funcs.values()), "不純が 0 ── 測定器が壊れている疑い"
    held = [k for k, v in funcs.items() if v["held"]]
    assert held, "差し替え名を読む関数が 0 ── 試験の setattr を拾えていない疑い"
    assert all(not funcs[k]["pure"] for k in held), "差し替え名を読むのに純と数えた関数がある"


def test_calling_something_that_stays_keeps_you_in():
    """★ 推移: 本体に残る関数を呼ぶ関数は、純と数えない（初版は依存を並べるだけで辿らなかった）。"""
    funcs = spc.survey()["funcs"]
    stays = {k for k, v in funcs.items() if not v["pure"]}
    wrong = [k for k, v in funcs.items() if v["pure"] and set(v["calls"]) & stays]
    assert not wrong, f"残る関数を呼ぶのに純と数えた: {wrong[:10]}"
    assert stays, "残る関数が 0（上の検査が空回りする）"
