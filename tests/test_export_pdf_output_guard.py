# データの出入口の盲検・致命4 と 高7（2026-08-26）── export-pdf の出力先まわり。
#
# 致命4: `--outdir <出力先の親>` を soffice に渡していたので、**必ず一度
#   `<ブック名>.pdf` を出力先フォルダに作る**。`--out` で別名を指定していても、
#   そこに同名の PDF が在れば予告なく消える。
#   実測: 顧客へ送った確定版 `請求書.pdf` が exit 0 のまま消滅。
#
# 高7: `--out` に上書きの関所が無い（export-csv には在る）。
#   ★ 関所の docstring は「export-pdf には --out が在ったので非対称でもあった」と
#     **自覚を書きながら**、直したのは --out が無かった側だけだった。
#
# 契約:
#   ① 出力先フォルダに現れるファイルは out_path ただ 1 つ（soffice に人のフォルダを触らせない）
#   ② 既存ファイルは --overwrite が無ければ上書きしない（exit 7・csv と同じ形と出口）
#   ③ 関所の実装は 1 つ（csv も pdf も同じ器官を通る）

import argparse
import inspect
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


# --- ③ 関所は 1 実装 ------------------------------------------------------------------

def test_both_exporters_go_through_one_gate():
    csv_src = inspect.getsource(ailine._export_csv_out_path)
    assert "_export_out_path(" in csv_src, "csv が共通の器官を通っていない"
    pdf_src = inspect.getsource(ailine.cmd_export_pdf)
    assert "_export_out_path(" in pdf_src, "pdf が共通の器官を通っていない"


# --- ② 上書きの関所 -------------------------------------------------------------------

def test_gate_refuses_an_existing_file(tmp_path):
    keep = tmp_path / "keep.pdf"
    keep.write_bytes(b"customer-final")
    a = argparse.Namespace(out=str(keep), overwrite=False)
    out_path, refuse = ailine._export_out_path(a, tmp_path / "default.pdf")
    assert refuse and "既にあります" in refuse, refuse
    assert keep.read_bytes() == b"customer-final"


def test_gate_lets_overwrite_through_when_asked(tmp_path):
    keep = tmp_path / "keep.pdf"
    keep.write_bytes(b"x")
    a = argparse.Namespace(out=str(keep), overwrite=True)
    _out, refuse = ailine._export_out_path(a, tmp_path / "default.pdf")
    assert refuse is None


def test_gate_is_silent_for_a_fresh_name(tmp_path):
    a = argparse.Namespace(out=None, overwrite=False)
    out_path, refuse = ailine._export_out_path(a, tmp_path / "new.pdf")
    assert refuse is None and out_path == tmp_path / "new.pdf"


def test_export_pdf_accepts_the_overwrite_flag():
    """② 関所を足したなら、通す手段も同じコマンドに在ること（行き止まりを作らない）。"""
    parser = ailine.build_parser() if hasattr(ailine, "build_parser") else None
    if parser is None:
        pytest.skip("パーサの入口の名前が変わった（契約を読み直すこと）")
    a = parser.parse_args(["export-pdf", "b.xlsx", "--overwrite"])
    assert a.overwrite is True


# --- ① soffice に人のフォルダを触らせない ---------------------------------------------

def test_soffice_is_never_pointed_at_the_users_folder(tmp_path, monkeypatch):
    """★ 致命4 の芯。`--outdir` に渡すのは**専用の一時フォルダ**であること。

    ★ 「同名 PDF が消えなかった」を状態で測ると、たまたま消えなかった実装を通す。
      ここでは **soffice に渡した引数そのもの**を見る（経路の不在を直接主張する）。
    """
    book = tmp_path / "請求書.xlsx"
    book.write_bytes(b"x")
    dst = tmp_path / "dst"
    dst.mkdir()
    victim = dst / "請求書.pdf"
    victim.write_bytes(b"customer-final")
    seen = {}

    class _Proc:
        returncode = 0
        stdout = ""
        stderr = ""

    def fake_run(cmd, **kw):
        seen["cmd"] = cmd
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        (outdir / (book.stem + ".pdf")).write_bytes(b"NEW-PDF")
        return _Proc()
    monkeypatch.setattr(ailine.subprocess, "run", fake_run)
    monkeypatch.setattr(ailine, "find_office_dir", lambda: Path(sys.executable).parent,
                         raising=False)
    ok, why = ailine._soffice_to_pdf(book, dst / "2026年1月分.pdf")
    # ★ 2026-08-27（CI が赤くなって気づいた）: 手元には LibreOffice が在るので通っていたが、
    #   CI には無い ── この repo が 5 度目に踏んだ「居るから見えない」。
    #   ★ skip の条件が甘かった: 「soffice」を含む理由だけを見ていたが、実際の断りは
    #     「LibreOffice の場所が分からない（basrun.py が見つかりません）」だった。
    #   ★ **測れない環境では skip と言う**（skip は「守っている」ではない ── だから
    #     この試験は -m local 側にも同じ契約を置かず、代わりに下の静的な番人で
    #     「人のフォルダを --outdir に渡していないこと」をコードから直接縛る）。
    if not ok and any(w in (why or "") for w in ("soffice", "LibreOffice", "basrun")):
        pytest.skip(f"この環境では LibreOffice を差し替えられない: {why}")
    assert ok, why
    outdir = Path(seen["cmd"][seen["cmd"].index("--outdir") + 1])
    assert outdir != dst, "人のフォルダを soffice の出力先にした（同名 PDF が消える）"
    assert victim.read_bytes() == b"customer-final", "名指ししていない PDF を消した"
    assert (dst / "2026年1月分.pdf").read_bytes() == b"NEW-PDF"
    assert sorted(q.name for q in dst.iterdir()) == ["2026年1月分.pdf", "請求書.pdf"]


def test_the_user_folder_is_never_the_outdir_by_construction():
    """★ 実行できない環境でも守れる形（上の試験は LibreOffice が要る）。

    「人のフォルダを soffice の出力先にしない」を**コードから直接**縛る ──
    実測で踏んだ壊れ方は `--outdir <出力先の親>` を渡していたことだった。
    """
    import inspect
    src = inspect.getsource(ailine._soffice_to_pdf)
    code = chr(10).join(ln.split("#")[0] for ln in src.split(chr(10)))
    assert "TemporaryDirectory" in code, "専用の一時フォルダを作っていない"
    assert "out_dir = Path(_pdf_tmp.name)" in code, (
        "出力先の親を --outdir に渡している ── 同名 PDF が予告なく消える")
