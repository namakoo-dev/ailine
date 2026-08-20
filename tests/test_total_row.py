"""単位L `ailine_core/total_row.py`（合計行の識別と、除外を算術で閉じる検査）の検体。
   ★ 実装前に書いた赤い検体（DESIGN-20260821-multifile v2.1・specimen-first）。

   背景（architect レビュー C3・2026-08-21 実測）: 既存 sum_identity は
   請求書形（小計+消費税+合計）の合計と、部署別形の総計を取り逃がし、
   ただのデータ行 300 を偽陽性で拾う。読み側（縦積み）は取り逃がし=黙って二重計上
   なので、書き側と非対称が反転する ── 広く候補を拾い、算術で閉じる。

   v2.1 の契約:
   - 排除トリガ = ラベル語（合計/計/小計/総計）∨ ラベル空白+数値 ∨ 直上に空行。語と構造のみ
   - 算術恒等はトリガでなく「閉じる検査」: 各除外行の値 = その行より上の採用行の和。
     不一致は 両側の数字つき で mismatches に出す（黙らない）
   - ★ sum_identity.py には触らない（「語を読まない」不変条件・番人つき）"""
from pathlib import Path
import sys

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from ailine_core.total_row import split_total_rows  # noqa: E402


def _rows(*pairs):
    """(ラベル, 値) の列を 1 行目からの行番号つきに展開する。"""
    return [(i + 1, label, value) for i, (label, value) in enumerate(pairs)]


def test_a_two_details_and_a_total_row_is_excluded_and_closes(tmp_path=None):
    """(a) 明細2+合計: ラベル語で除外・閉じる検査が通る。"""
    v = split_total_rows(_rows(("品A", 100.0), ("品B", 200.0), ("合計", 300.0)))
    assert [r.row for r in v.excluded] == [3]
    assert v.adopted_rows == [1, 2]
    assert v.mismatches == []


def test_b_single_detail_and_total_is_still_excluded(tmp_path=None):
    """(b) 明細1+合計: ★ 旧 sum_identity は MIN_TERMS=2 で沈黙した形。ラベル語で拾う。"""
    v = split_total_rows(_rows(("品A", 100.0), ("合計", 100.0)))
    assert [r.row for r in v.excluded] == [2]
    assert v.mismatches == []


def test_c_invoice_shape_subtotal_tax_total_all_close(tmp_path=None):
    """(c) 請求書形: 小計と合計が除外・消費税は採用に残る・両方の除外が算術で閉じる。
       合計 1100 = 明細 1000 + 税 100（Σ保存 ── 縦積みの合計金額が狂わない形）。"""
    v = split_total_rows(_rows(
        ("品A", 600.0), ("品B", 400.0), ("小計", 1000.0), ("消費税", 100.0), ("合計", 1100.0)))
    assert [r.row for r in v.excluded] == [3, 5]
    assert v.adopted_rows == [1, 2, 4], "消費税はラベル語でも構造でもないので採用に残る"
    assert v.mismatches == []


def test_d_department_shape_with_coincidence_row(tmp_path=None):
    """(d) 部署別形: 小計・総計は除外して閉じる。★ 偶然 上の和に等しいだけのデータ行
       （100+200 の後の 300・C3 の実測偽陽性）は採用のまま。"""
    v = split_total_rows(_rows(
        ("部署A", 100.0), ("部署B", 200.0), ("小計", 300.0),
        ("部署C", 300.0), ("総計", 600.0)))
    assert [r.row for r in v.excluded] == [3, 5]
    assert 4 in v.adopted_rows, "偶然の一致行を偽陽性で除外している"
    assert v.mismatches == []


def test_e_total_row_with_blank_label_is_caught_by_structure(tmp_path=None):
    """(e) ラベルが空の合計行: 語が無くても『ラベル空白+数値』の構造トリガで拾う。
       ★ 根 1（A 列空白で表が切れる）と単位L の効果を分離する校正検体。"""
    v = split_total_rows(_rows(("品A", 100.0), ("品B", 200.0), (None, 300.0)))
    assert [r.row for r in v.excluded] == [3]
    assert v.mismatches == []


def test_f_mismatched_total_is_excluded_with_both_numbers(tmp_path=None):
    """(f) 合計の値が合わない: 除外はするが、閉じる検査が 両側の数字つき で鳴る
       （「一致しませんでした」だけの報告は感想 ── 信用の条件④）。"""
    v = split_total_rows(_rows(("品A", 600.0), ("品B", 400.0), ("合計", 1200.0)))
    assert [r.row for r in v.excluded] == [3]
    assert len(v.mismatches) == 1
    m = v.mismatches[0]
    assert m.row == 3 and m.excluded_value == 1200.0 and m.adopted_sum == 1000.0


def test_g_no_total_rows_noise_floor(tmp_path=None):
    """(g) ノイズ床（V3）: 合計行が無い普通の表では何も除外しない・何も鳴らない。"""
    v = split_total_rows(_rows(("品A", 100.0), ("品B", 200.0), ("品C", 50.0)))
    assert v.excluded == [] and v.mismatches == []
    assert v.adopted_rows == [1, 2, 3]


def test_h_blank_row_above_triggers_structure(tmp_path=None):
    """(h) 直上に空行がある数値行（罫線代わりの空行の下の Total 行）: ラベル語に無い
       英語 Total でも構造トリガで拾い、閉じる検査で正しさを確かめる。"""
    v = split_total_rows(_rows(
        ("部署A", 100.0), ("部署B", 200.0), (None, None), ("Total", 300.0)))
    assert [r.row for r in v.excluded] == [4]
    assert v.mismatches == []


def test_i_fragment_label_like_sekkeibu_is_not_a_trigger(tmp_path=None):
    """(i) ★ 断片ガード: 『設計部』は『計』を含むがラベル語トリガにしない
       （『計』は完全一致のみ・2 文字語 合計/小計/総計 は包含で可）。
       値が偶然 上の和と等しくても（100+200 の後の 300）採用のまま。"""
    v = split_total_rows(_rows(
        ("部署A", 100.0), ("部署B", 200.0), ("設計部", 300.0), ("合計", 600.0)))
    assert [r.row for r in v.excluded] == [4]
    assert 3 in v.adopted_rows, "『設計部』を『計』の部分一致で誤爆している"
    assert v.mismatches == []


def test_module_does_not_import_sum_identity_mutation(tmp_path=None):
    """★ 番人: total_row は sum_identity を「呼ぶ側」に立つ（触らない・上書きしない）。
       sum_identity の公開関数がこの import で書き換わっていないこと。"""
    import ailine_core.sum_identity as si
    import importlib
    importlib.reload(si)
    assert hasattr(si, "rows_matching_sum_above")
