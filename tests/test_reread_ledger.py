# 読み直しの層の**台帳**（2026-09-02）── 15 塊のうち、何が守られているか。
#
# ★★ 2026-09-02 に実際に踏んだ形（「〜以外」の実装中）:
#   読み直しが正しく `cmp=nin` を立てたのに、**決定の場所（verify_dsl_args）が
#   否定を知らないまま `eq` に上書き**し、「味噌汁以外を抜き出して」で
#   **味噌汁だけを抜き出して △** が出た ── 逆のことをして合格。
#   ★ 読み直しに足して、**決定の場所に足し忘れた**。片配線そのもの。
#
# ★★ 数えて分かったこと:
#   読み直しの塊は **15 個**。うち「読み直しました」と**文言を出すのは 4 つだけ**で、
#   残り 11 は**黙って読み替えている**（解釈行には出るが、読み替えたこと自体は言わない）。
#   ★ 黙って読み替えるのが悪いとは限らない ── だが「どれが黙るか」を**誰も数えていなかった**。
#
# ★ この台帳は「守られているか」を断定しない。**数えるだけ**にする ──
#   1 塊ずつ検体を書いて減らしていく作業の、分母を出すのが仕事。
#
# 契約:
#   ① 読み直しの塊の数が変わったら気づく（増えた塊は無防備で入ってくる）
#   ② 文言を出す塊の数が減ったら気づく（黙る方向への変化は、伝わらなくなる変化）
#   ③ 配線を通す検体を持つ塊を、名指しで台帳に持つ

import inspect
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

SRC = inspect.getsource(ailine)


def _reread_segment() -> list:
    lines = SRC.splitlines()
    i0 = next(i for i, l in enumerate(lines) if l.startswith("def _translate_and_dispatch("))
    i1 = next(i for i in range(i0 + 1, len(lines)) if lines[i].startswith("def "))
    return lines[i0:i1]


def _blocks() -> list:
    """(印を立てる行, その塊が出す文言 or None) の一覧。"""
    out, msg = [], None
    for i, ln in enumerate(_reread_segment()):
        m = re.search(r"『(.+?)』として読み直しました", ln)
        if m:
            msg = m.group(1)
        if "_reread_done = True" in ln:
            out.append((i, msg))
            msg = None
    return out


# --- 2026-09-02 の初回計測 ---------------------------------------------------------
BLOCKS_AT_FIRST_COUNT = 15
SPEAKING_AT_FIRST_COUNT = 4          # 「〜として読み直しました」と言う塊

# --- 配線を通す検体を持つ塊（★ 名指しで持つ・増やしていく）------------------------------
#   ここに書けるのは「LLM を差し替えて読み直しの経路を通し、**決定の場所を抜けた後の
#   結果**まで見ている」検体だけ。解釈行を見るだけのものは含めない。
COVERED = {
    "セル 2 つの入れ替え": "tests/test_swap_two_cells.py::"
                            "test_the_cell_reread_actually_fires_without_the_llm",
    "行の入れ替え": "tests/test_swap_rows_resolved_by_machine.py::"
                     "test_the_reread_actually_fires_without_the_llm",
    "『〜以外』の抽出": "tests/test_except_extraction.py::"
                         "test_the_reread_fires_through_the_real_path",
    # --- 第 2 波（2026-09-05・盲検の査定が 681 行を指摘したので、畳む前に分母を減らす）---
    "数値書式": "tests/test_reread_specimens_wave2.py::"
                 "test_the_number_format_reread_fires_without_the_llm",
    "列抽出": "tests/test_reread_specimens_wave2.py::"
               "test_the_column_extraction_reread_fires_without_the_llm",
    "列追加": "tests/test_reread_specimens_wave2.py::"
               "test_the_add_column_reread_uses_the_second_pass",
    "置き換え": "tests/test_reread_specimens_wave2.py::"
                 "test_the_replace_reread_does_not_overwrite_the_whole_column",
    "条件つき書換": "tests/test_reread_specimens_wave2.py::"
                     "test_the_conditional_write_reread_fires_on_the_split_plan",
}


def test_the_number_of_reread_blocks_is_watched():
    """① 塊の数 ── 増えたら、無防備な読み直しが 1 つ入ったということ。"""
    n = len(_blocks())
    assert n == BLOCKS_AT_FIRST_COUNT, (
        f"読み直しの塊が {n} 個（初回計測は {BLOCKS_AT_FIRST_COUNT} 個）── "
        "増やしたなら、その塊の配線を通す検体を書いて COVERED に足し、この数も更新すること")


def test_the_number_of_speaking_blocks_does_not_shrink():
    """② 黙る方向への変化に気づく。

    ★ 「読み直しました」を言わない塊が増えると、**読み替えたこと自体が伝わらない**。
      この repo は 8/31 に「判定は正しいが、説明の文面が古い」で 1 日潰している。
    """
    speaking = sum(1 for _i, msg in _blocks() if msg)
    assert speaking >= SPEAKING_AT_FIRST_COUNT, (
        f"文言を出す塊が {speaking} 個に減った（初回計測は {SPEAKING_AT_FIRST_COUNT} 個）")


def test_the_covered_specimens_exist():
    """③ 台帳に書いた検体が、実在すること（名前だけの安心を作らない）。"""
    assert COVERED, "配線を通す検体の表が空（★ 空だとこの番人は 1 度も回らず黙る）"
    for label, ref in COVERED.items():
        f, _, name = ref.partition("::")
        path = REPO / f
        assert path.exists(), f"{label}: {f} が無い"
        assert name in path.read_text(encoding="utf-8"), f"{label}: {name} が無い"


def test_the_backlog_is_visible():
    """★ 分母を出す ── 15 塊のうち、配線を通す検体を持つのは 8 塊（第 2 波で 3 → 8）。

    ★ これは「12 塊が壊れている」という意味ではない。**測っていない**という意味。
      減らしていくための数として置く（増えたら赤くする）。
    """
    n = len(_blocks())
    assert len(COVERED) <= n
    assert len(COVERED) >= 8, f"配線を通す検体が減っている（{len(COVERED)} 件）"
