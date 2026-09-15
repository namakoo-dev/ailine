# 復元の致命2（2026-08-24 の盲検）── 隣にある利用者のファイルを無言で消す。
#
# ★ 実測: `<book>.out.xlsx` は作業ファイル名として**存在確認も警告もなく**上書きされ、
#   原本反映が成功すると unlink される。利用者が自分で作った `売上.out.xlsx` は
#   一言も無く消えた（`--copy` の時は中身だけ上書きされ、しかも
#   「（原本 売上.xlsx は変更していません）」と表示される ── 別の原本は破壊済み）。
# ★ フォルダ経路には同じ危険への関所（_refuse_output_conflict・exit 7・
#   「ailine の印が無い人のファイルです」）が**既に在る**のに、単一ブック経路に
#   配線されていなかった ── 片配線。
#
# 契約:
#   ① 人のファイルが出力先に在れば、**触る前に**止める（exit 7）
#   ② ailine 産（前回の .out）なら従来どおり黙って作り直す
#   ③ 出力先が空いていれば 1 文字も増えない
#   ④ `.out` の場所を決める実装は 1 つ（4 箇所の書き写しを畳む）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402
from test_golden_transcripts import _isolate, _run_main  # noqa: E402
from _product_source import count_in_product, product_text  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


def _book(tmp_path, name="売上.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"]); ws.append(["a", 100]); ws.append(["b", 250])
    wb.save(p)
    return p


def test_out_path_has_one_implementation():
    """④ 4 箇所が同じ形を書き写していた ── 畳んだことを構造で縛る。"""
    assert count_in_product('with_name(book.stem + ".out"') <= 1, \
        ".out の場所を決める式が複数ある（書き写し）"
    assert "def out_book_path(" in product_text(), "共通の実装が無い"


def test_refuses_when_a_human_file_sits_at_the_output_path(tmp_path, monkeypatch, capsys):
    """① 人のファイルは触る前に守る。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    mine = tmp_path / "売上.out.xlsx"
    wb = openpyxl.Workbook(); wb.active["A1"] = "私の大事なメモ"; wb.save(mine)
    before = mine.read_bytes()
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    rc, out = _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)
    assert rc == 7, f"人のファイルを守らなかった: exit={rc} / {out}"
    assert mine.read_bytes() == before, "人のファイルが書き換わった"
    assert "人のファイル" in out or "書けません" in out, out


def test_own_previous_output_is_rebuilt_silently(tmp_path, monkeypatch, capsys):
    """② 前回の .out（ailine 産）は従来どおり作り直す ── 誤爆で使えなくしない。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        ws = wb.active
        ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "b", 250, "a", 100
        wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    rc1, _ = _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)
    assert rc1 == 0
    rc2, out2 = _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)
    assert rc2 == 0, f"自分の前回出力で止まった（誤爆）: {out2}"


# ── ★ 2026-09-16: 関所で断った run が行き止まりを作らないこと（盲検の買い手役②）────────
#
# ★★ 事故（再現済み）:
#   1 回目 集計 → ✓ ／ 2 回目 別の集計 → 上書きの関所 exit 7。このとき `.out.xlsx` が
#   **黙って**残り、3 回目以降は出力先の関所が
#   「この道具が書いた記録がありません（人が置いたファイルか…）」で全部 exit 7。
#   **1 分前にこの道具自身が作った物**なのに。`run` に上書き許可のフラグが無いので、
#   人がファイルを消すまでその本には二度と実行できない ── 買い手の言葉で
#   「事務職に配れる品質ではありません」。
# ★ 原因: `_finish_run` は `result["out"]` が在れば指紋を履歴に残し、次の run が
#   「俺が置いたまま」と分かる。**関所で断つ 4 経路だけ** `_finish_run` を通っていなかった。
#   ★ 2026-08-26 に `_finish_failed_apply` で直したのと同じ形（あれは「反映に失敗」側だけ）。

def _agg(col):
    return {"op": "AGGREGATE", "args": {"group_col": col, "value_col": "金額"}}


def _book2(tmp_path):
    p = tmp_path / "売上2.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上2"
    for row in [["取引先", "金額", "担当"], ["a", 100, "甲"], ["b", 250, "乙"], ["a", 50, "甲"]]:
        ws.append(row)
    wb.save(p)
    return p


#: ★ 治具がどの列で集計するか ── 検体が宣言する（`code` は列を番号で書くので読めない）。
_FAKE_GROUP = ["取引先"]


def _fake_agg(out_book, code, workdir, helper_files=(), timeout=None):
    """実 LO の代わり: 頼まれた列で本当に集計して『集計』シートを作る。

    ★ 初版は常に取引先で集計していたので、2 回目（担当ごと）が**関所の手前で**
      事後条件に落ち、測りたい所まで届かなかった ── 治具も検体の一部（窒息点を仮説する）。
    """
    group = _FAKE_GROUP[0]     # ★ 検体が宣言する（Basic は列を番号で書くので code からは読めない）
    wb = openpyxl.load_workbook(out_book)
    src = wb["売上2"]
    headers = [c.value for c in src[1]]
    gi, vi = headers.index(group), headers.index("金額")
    totals = {}
    for row in src.iter_rows(min_row=2, values_only=True):
        if row[gi] is None:
            continue
        totals[row[gi]] = totals.get(row[gi], 0) + (row[vi] or 0)
    if "集計" in wb.sheetnames:
        del wb["集計"]
    sh = wb.create_sheet("集計")
    sh.append([group, "合計 - 金額"])
    for k, v in totals.items():
        sh.append([k, v])
    wb.save(out_book)
    return True, None, "ok"


def test_a_gated_run_says_it_left_a_file_and_does_not_wall_the_book(tmp_path, monkeypatch, capsys):
    """★★ 関所で断った回が ① 残したことを言い ② 次の run を塞がないこと。

    ★ 見分けは**両方**で取る ── 言うだけ直して記録を忘れると行き止まりは残り、
      記録だけ直して黙ると人は隣のファイルに気づかない（片方だけ直さない）。
    """
    _isolate(monkeypatch, tmp_path)
    book = _book2(tmp_path)
    monkeypatch.setattr(ailine, "basrun_apply", _fake_agg)
    _FAKE_GROUP[0] = "取引先"      # ★ 検体ごとに宣言する（モジュール変数の持ち越しを断つ）

    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: _agg("取引先"))
    rc1, _o1 = _run_main(["run", str(book), "取引先ごとに金額を合計して", "--overwrite"], capsys)
    assert rc1 == 0, _o1

    # 2 回目: 既存の『集計』を書き換えるので上書きの関所が立つ（端末が無いので exit 7）
    _FAKE_GROUP[0] = "担当"
    # ★ 端末が無い場を作る（stdin を触らず、道具が用意している差し替え口を使う）
    monkeypatch.setattr(ailine, "_stdin_isatty", lambda: False)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: _agg("担当"))
    rc2, o2 = _run_main(["run", str(book), "担当ごとに金額を合計して"], capsys)
    assert rc2 == 7, o2
    leftover = tmp_path / "売上2.out.xlsx"
    if leftover.exists():
        assert "作業結果は" in o2 and leftover.name in o2, (
            f"関所で断ったのに、残した物を言っていない\n{o2}")

    # 3 回目: ★ ここが行き止まりだった
    _FAKE_GROUP[0] = "取引先"
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: _agg("取引先"))
    rc3, o3 = _run_main(["run", str(book), "取引先ごとに金額を合計して", "--overwrite"], capsys)
    assert rc3 == 0, (f"関所で断った回の残骸が、次の run を塞いでいる（行き止まり）\n{o3}")
    assert "この道具が書いた記録がありません" not in o3, o3


def test_a_file_the_person_edited_is_still_refused(tmp_path, monkeypatch, capsys):
    """★ 陰性対照 ── 自分の物でも**人が手を入れた**後なら、従来どおり断る。
       （行き止まりを消すために、人の作業を踏み潰す側へ倒していないこと）"""
    _isolate(monkeypatch, tmp_path)
    book = _book2(tmp_path)
    monkeypatch.setattr(ailine, "basrun_apply", _fake_agg)
    _FAKE_GROUP[0] = "取引先"      # ★ 同上
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: _agg("取引先"))
    rc1, o1 = _run_main(["run", str(book), "取引先ごとに金額を合計して", "--copy"], capsys)
    assert rc1 == 0, o1
    out = tmp_path / "売上2.out.xlsx"
    assert out.exists(), o1
    wb = openpyxl.load_workbook(out)
    wb.active["Z99"] = "人が書いた"
    wb.save(out)

    rc2, o2 = _run_main(["run", str(book), "取引先ごとに金額を合計して", "--copy"], capsys)
    assert rc2 == 7, f"人が手を入れた出力を黙って上書きした\n{o2}"
    assert "そのあと変更されています" in o2, o2
