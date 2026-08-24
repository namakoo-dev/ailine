# フォルダの分母 ── 実装より先に凍結した赤い検体（2026-08-24）。
#
# ★ 根（盲検 3 者が別々の入口から同じ形に着いた）:
#   **処理できたものだけを数えて「元」と呼び、それと出力を比べている。**
#   だから処理できなかったものは元側にも現れず、比較が恒真になる。
#
# ★ 実測（2 者が独立に再現）: `.xlsm` を混ぜたフォルダで
#     「3 ファイル中 3 照合できた」「Σ金額 元 4500 / 出力 4500 ✓」
#   と出る。**6 冊のうち 3 冊が無かったことになっている。**
#   マクロ入りの請求書テンプレは、実際の経理フォルダで最も在りうる非 .xlsx。
#
# 契約:
#   ① フォルダに在るのに候補にしなかったファイルは、**除外として数えて名指しできる**
#   ② 既に手当てされている分類（~$ ロック・サブフォルダ・.csv）は変えない
#   ③ ★ 恒真殺し: .xlsm を 1 つ増やしたら、除外の件数が 1 増えること

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core.multifile import classify_folder_contents  # noqa: E402


def _make(tmp_path, names):
    for n in names:
        p = tmp_path / n
        if n.endswith("/"):
            (tmp_path / n[:-1]).mkdir()
        else:
            p.write_bytes(b"x")
    return tmp_path


def test_unreadable_formats_are_counted_not_dropped(tmp_path):
    """① .xlsm / .xlsb / .ods が分母から消えない。"""
    f = _make(tmp_path, ["a.xlsx", "b.xlsx", "c.xlsm", "d.xlsb", "e.ods"])
    candidates, excluded = classify_folder_contents(f)
    assert [p.name for p in candidates] == ["a.xlsx", "b.xlsx"]
    assert excluded.get("other_format", 0) == 3, \
        f"3 冊が分母から消えた（除外に数えていない）: {excluded}"


def test_the_excluded_names_can_be_shown(tmp_path):
    """★ 件数だけでは人は動けない。**どのファイルか**を名指しできること。"""
    f = _make(tmp_path, ["a.xlsx", "請求書テンプレ.xlsm"])
    _c, excluded = classify_folder_contents(f)
    names = excluded.get("other_format_names") or []
    assert "請求書テンプレ.xlsm" in names, f"名指しできない: {excluded}"


def test_adding_one_xlsm_increases_the_count_by_one(tmp_path):
    """③ 恒真殺し: 増やしたら増えること（数えているふりでないこと）。"""
    f1 = _make(tmp_path / "one", ["a.xlsx", "x.xlsm"]) if (tmp_path / "one").mkdir() is None else None
    f2 = tmp_path / "two"; f2.mkdir()
    _make(f2, ["a.xlsx", "x.xlsm", "y.xlsm"])
    _c1, e1 = classify_folder_contents(tmp_path / "one")
    _c2, e2 = classify_folder_contents(f2)
    assert e2["other_format"] == e1["other_format"] + 1, (e1, e2)


def test_existing_buckets_are_unchanged(tmp_path):
    """② 既存の分類（~$ / サブフォルダ / .csv）は変えない。"""
    f = tmp_path / "mix"; f.mkdir()
    _make(f, ["a.xlsx", "~$a.xlsx", "b.csv", "sub/"])
    candidates, excluded = classify_folder_contents(f)
    assert [p.name for p in candidates] == ["a.xlsx"]
    assert excluded["temp"] == 1 and excluded["csv"] == 1 and excluded["subdirs"] == 1


# --- ★ 積めなかった冊があるのに exit 0（2026-08-24・盲検の使い勝手レビュー）------------
#
# 実測: 列名の違う冊を名指しで断りながら **exit 0**。しかも Σ の「元」は積めた冊だけの和
# なので**必ず一致する（恒真）**。得意先 1 社 283,500 円が消えても、
# スクリプトからは「成功・Σ 一致」にしか見えなかった。
#
# 契約: 積めなかった冊 / 読めない形式が 1 つでもあるなら **exit 0 にしない**
#       （出力は作る ── 人が続きを決められるように）。

import subprocess as _sp


def _run_stack(folder, out):
    import sys as _s
    return _sp.run([_s.executable, "-m", "ailine", "stack", str(folder), "--out", str(out)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace",
                    cwd=str(Path(__file__).resolve().parent.parent),
                    env={**__import__("os").environ, "PYTHONPATH": "src"})


def test_stack_does_not_exit_zero_when_a_book_could_not_be_stacked(tmp_path):
    import openpyxl
    f = tmp_path / "s"; f.mkdir()
    for n, hdr in (("a.xlsx", "金額"), ("b.xlsx", "金額"), ("c.xlsx", "金額(税抜)")):
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["商品", hdr]); ws.append(["x", 100])
        wb.save(f / n)
    r = _run_stack(f, tmp_path / "out.xlsx")
    assert "c.xlsx" in r.stdout, r.stdout
    assert r.returncode != 0, f"積めなかった冊が在るのに exit 0（{r.stdout}）"
    assert (tmp_path / "out.xlsx").exists(), "出力自体は作ってよい（人が続きを決められる）"


def test_stack_exits_zero_when_everything_was_stacked(tmp_path):
    """誤爆防止: 全冊積めたら今までどおり exit 0。"""
    import openpyxl
    f = tmp_path / "ok"; f.mkdir()
    for n in ("a.xlsx", "b.xlsx"):
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["商品", "金額"]); ws.append(["x", 100])
        wb.save(f / n)
    r = _run_stack(f, tmp_path / "out2.xlsx")
    assert r.returncode == 0, r.stdout + r.stderr


def test_ailine_own_workfiles_are_not_counted_as_unreadable(tmp_path):
    """★ ailine 自身の作業ファイルを「読めない形式」に数えない（2026-08-24）。

    実測: `~/.ailine` をフォルダ内に隔離した検体で、`history.jsonl` が
    「対象外: 読めない形式 1 件（history.jsonl）── .xlsx に保存し直すと扱えます」
    と報告された。**自分が置いたものを他人の資料と同じに扱っている**。
    実運用でも、作業フォルダに ailine 産のファイルが同居すれば同じことが起きる。
    ★ 分母に入れるのは「人が置いた、扱えなかったもの」だけ。
    """
    f = tmp_path / "w"; f.mkdir()
    (f / "a.xlsx").write_bytes(b"x")
    for n in ("history.jsonl", "run.lock", "vocab.json", "aliases.json",
               "misclass.jsonl", "notice_v2_shown"):
        (f / n).write_bytes(b"{}")
    _c, excluded = classify_folder_contents(f)
    assert excluded.get("other_format", 0) == 0, \
        f"ailine 自身の作業ファイルを分母に数えた: {excluded.get('other_format_names')}"
