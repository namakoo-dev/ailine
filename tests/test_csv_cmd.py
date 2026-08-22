# CSV 接続（cmd_run_csv / `ailine csv`）の決裁済み不変条件 ── 実装より先に凍結した赤い検体
#
# 出典: DESIGN-20260821-multifile.md「CSV 検疫 設計 v2」+ REVIEW-20260822-csv-architect.md §8。
# ここで凍結するのは**決裁済みの線**だけ（文言の細部は実装側の設計に任せる）:
#   ① `ailine csv <file>` は <stem>.xlsx を隣に作り、原本 CSV は 1 バイトも変えない
#   ② ✓ の文言は転送段の宣言（欠落/不一致/余剰の 3 計数）であって、
#      「正しく読み込みました」とは**言わない**。文字列保持した列があるときは
#      「Σ で検算していません」系の開示を必ず出す
#   ③ .csv が normalize_book / basrun_apply に到達しない（LO の CSV インポートが
#      0 落ちの発生源 ── 実測 0123→123。構造の番人 = PC5 変異の常設形）
#   ④ 照合（2 冊）の入口に .csv を渡したとき、「xlsx 形式か確認してください」の
#      誤誘導で終わらず、csv の扱いへ誘導する

import hashlib
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "cmd_run_csv"),
    reason="CSV 接続 未実装（不変条件は凍結済み・実装が来たら自動で実測に切り替わる）",
    strict=True,
)


def _nasty_csv(tmp_path, name="仕入.csv"):
    # 0落ち・カンマ列・数式頭を含む cp932 の実物形
    raw = ("コード,金額,メモ\n"
           "0123,\"1,234\",=SUM(1:2)\n"
           "00456,\"2,345\",通常\n").encode("cp932")
    p = tmp_path / name
    p.write_bytes(raw)
    return p, raw


@needs_impl
def test_csv_command_transfers_without_claiming_read_correctness(tmp_path, monkeypatch, capsys):
    """①+②: 変換は成功し、主張は転送段の 3 計数のみ。原本は無変更。"""
    _isolate(monkeypatch, tmp_path)
    csv_path, raw = _nasty_csv(tmp_path)
    rc, out = _run_main(["csv", str(csv_path)], capsys)
    assert rc == 0, out
    out_x = csv_path.with_suffix(".xlsx")
    assert out_x.exists(), f"<stem>.xlsx が隣にできていない: {out}"
    # 原本 CSV は 1 バイトも変わらない
    assert hashlib.sha256(csv_path.read_bytes()).hexdigest() == \
        hashlib.sha256(raw).hexdigest()
    # 転送段の 3 計数（欠落/不一致/余剰）を必ず宣言する
    assert "欠落" in out and "不一致" in out and "余剰" in out, out
    # 「正しく読み込みました」とは言わない（言えるのは転送だけ）
    assert "正しく読み込みました" not in out, out
    # 文字列保持した列（コード: 先頭ゼロ）があるので Σ 対象外の開示が要る
    assert "Σ" in out and "検算していません" in out, out


@needs_impl
def test_csv_never_reaches_lo_or_normalize(tmp_path, monkeypatch, capsys):
    """③: 構造の番人。csv 経路で normalize_book / basrun_apply に .csv が渡らない
       （そもそも呼ばれないのが期待形 ── 呼ばれたら即座に落とす）。"""
    _isolate(monkeypatch, tmp_path)
    csv_path, _raw = _nasty_csv(tmp_path)

    def _trap_normalize(path, *a, **kw):
        raise AssertionError(f"normalize_book が csv 経路で呼ばれた: {path}")

    def _trap_basrun(*a, **kw):
        raise AssertionError("basrun_apply が csv 経路で呼ばれた")

    monkeypatch.setattr(ailine, "normalize_book", _trap_normalize)
    monkeypatch.setattr(ailine, "basrun_apply", _trap_basrun)
    rc, out = _run_main(["csv", str(csv_path)], capsys)
    assert rc == 0, out
    # 出力の 0落ち生存を読み戻しで確認（LO を通っていない証拠の実側）
    from ailine_core import xml_readback
    grid = xml_readback.read_grid(csv_path.with_suffix(".xlsx"))["grid"]
    assert grid[(2, 1)] == "0123", f"0落ち: {grid.get((2, 1))!r}"


@needs_impl
def test_match_entrance_redirects_csv_instead_of_misleading(tmp_path, monkeypatch, capsys):
    """④: 2 冊照合の入口に .csv ── 「xlsx 形式か確認してください」の誤誘導で
       終わらず、csv の扱い（ailine csv）へ誘導する。"""
    _isolate(monkeypatch, tmp_path)
    csv_path, _raw = _nasty_csv(tmp_path, name="a.csv")
    csv2, _ = _nasty_csv(tmp_path, name="b.csv")
    rc, out = _run_main(["run", str(csv_path), str(csv2), "2冊を照合して"], capsys)
    assert rc != 0
    assert "csv" in out.lower(), f"csv への言及が無い: {out}"
    assert "xlsx 形式か確認してください" not in out, f"誤誘導が残っている: {out}"
