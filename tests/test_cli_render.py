"""C8: ailine_core/cli_render.py のレンダラ関数の単体テスト。

   ★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。
   ★ 統合側（実際に ailine.py の各コマンドがこれらの関数を通して出す文言そのもの・
   byte 一致）は tests/golden/ 側（F9 端末トランスクリプト等）が引き続き見る
   （C8 は表示層の移動＋一部統合で、出力そのものは1バイトも変えていない＝
   golden 256本の差分ゼロが実証）。ここでは各レンダラ単体の入出力を検査する。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ailine_core.cli_render import (
    render_code_block, render_retry_options, render_aborted, render_run_header,
    render_backup_list, render_restore_done, render_vocab_add_result, render_vocab_listing,
)


# --- render_code_block -----------------------------------------------------

def test_render_code_block_no_step_prefix():
    lines = render_code_block("\n─ 生成した .bas（ルール変換・LLM不使用）───────────────", "Sub Run()\nEnd Sub")
    assert lines == [
        "\n─ 生成した .bas（ルール変換・LLM不使用）───────────────",
        "Sub Run()\nEnd Sub",
        "──────────────────────────────────────────",
    ]

def test_render_code_block_with_step_prefix():
    lines = render_code_block("  2段目: ─ 生成した .bas（語彙外・AI が直接作成）───────────────",
                               "code", step_prefix="  2段目: ")
    assert lines[-1] == "  2段目: ──────────────────────────────────────────"


# --- render_retry_options: 実測した3ゲートの桁揃えを個別に固定 ------------------

def test_render_retry_options_fidelity_gate_alignment():
    lines = render_retry_options("", [
        ("--accept-loss", "失われてよい（バックアップから ailine undo で復元可能）"),
        ("--copy", "原本には触らず .out に結果を作る（原本は無変更）"),
    ])
    assert lines[0] == "この処理を続けるには、以下のいずれかを指定して再実行してください:"
    assert lines[1] == "  --accept-loss  失われてよい（バックアップから ailine undo で復元可能）"
    assert lines[2] == "  --copy         原本には触らず .out に結果を作る（原本は無変更）"

def test_render_retry_options_overwrite_gate_alignment():
    lines = render_retry_options("  2段目: ", [
        ("--overwrite", "上書きを承知して続行する（バックアップから ailine undo で戻せる）"),
        ("--copy", "原本には触らず .out に結果を作る（原本は無変更）"),
    ])
    assert lines[1] == "  2段目:   --overwrite  上書きを承知して続行する（バックアップから ailine undo で戻せる）"
    assert lines[2] == "  2段目:   --copy       原本には触らず .out に結果を作る（原本は無変更）"

def test_render_retry_options_single_option_alignment():
    lines = render_retry_options("", [
        ("--allow-freeform", "機械検証できないことを承知の上で適用する"),
    ])
    assert lines[1] == "  --allow-freeform  機械検証できないことを承知の上で適用する"


# --- render_aborted ---------------------------------------------------------

def test_render_aborted_no_prefix():
    assert render_aborted() == "× 中止した"

def test_render_aborted_with_step_prefix():
    assert render_aborted("  3段目: ") == "  3段目: × 中止した"


# --- render_run_header ------------------------------------------------------

def test_render_run_header():
    assert (render_run_header("DSL 経路", "qwen2.5-coder:7b", "book.xlsx")
            == "■ ailine（DSL 経路）  model=qwen2.5-coder:7b  book=book.xlsx")

def test_render_run_header_plan_label_carries_count():
    assert (render_run_header("複合計画・3 段", "m", "b.xlsx")
            == "■ ailine（複合計画・3 段）  model=m  book=b.xlsx")


# --- render_backup_list -----------------------------------------------------

def test_render_backup_list_empty():
    assert render_backup_list("book.xlsx", []) == ["book.xlsx のバックアップは無い"]

def test_render_backup_list_nonempty():
    backups = [Path("book.xlsx.bak.2"), Path("book.xlsx.bak.1")]
    lines = render_backup_list("book.xlsx", backups)
    assert lines[0] == "book.xlsx のバックアップ（2 世代・新しい順）:"
    assert lines[1:] == ["book.xlsx.bak.2", "book.xlsx.bak.1"]


# --- render_restore_done ----------------------------------------------------

def test_render_restore_done_restore_no_remaining():
    assert (render_restore_done("book.xlsx", "book.xlsx.bak.1")
            == "✓ book.xlsx を book.xlsx.bak.1 から復元した")

def test_render_restore_done_undo_with_remaining():
    assert (render_restore_done("book.xlsx", "book.xlsx.bak.1", remaining=3)
            == "✓ book.xlsx を book.xlsx.bak.1 から復元した（あと 3 回戻せます）")


# --- vocab -------------------------------------------------------------------

def test_render_vocab_add_result_ok():
    assert render_vocab_add_result(True, "税率 = 0.08 を登録した") == "✓ 税率 = 0.08 を登録した"

def test_render_vocab_add_result_fail():
    assert render_vocab_add_result(False, "値が数値でない") == "× 値が数値でない"

def test_render_vocab_listing_empty():
    vocab_file = Path("dummy") / "vocab.json"
    lines = render_vocab_listing({}, vocab_file)
    assert lines == [f"（用語集は空。{vocab_file} に登録するか"
                      " `ailine vocab add <語> <値>` で追加）"]

def test_render_vocab_listing_sorted_and_formatted():
    vocab_file = Path("dummy") / "vocab.json"
    vocab = {"税込": 1.1, "税抜": 0.9}
    lines = render_vocab_listing(vocab, vocab_file)
    assert lines[0] == f"用語集（{vocab_file}・2件）:"
    assert lines[1:] == ["  税抜 = 0.9", "  税込 = 1.1"]
