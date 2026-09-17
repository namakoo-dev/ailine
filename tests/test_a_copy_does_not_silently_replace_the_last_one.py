# -*- coding: utf-8 -*-
"""`--copy` の成果を黙って置き換えない／別の名前に出せる（2026-09-17・盲検 3 体目）。

★★ 起きたこと（製造業の購買担当・初見・盲検）:
  `--copy` は必ず `<book>.out.xlsx` に書く。1 回目に仕入先別の集計を作り、2 回目に
  抽出を打ったら、**1 回目の『集計』シートは消えていた。警告は 1 行も無い。**
  買い手の言葉:「**3 つの作業をしたら 2 つ失います**」。
  逃げ道として `--out` を**在ると思って打ち**、「× 知らない指定があります: --out」で
  行き止まりになった。

★★ この損は**いちばん用心深い人に当たる** ── 原本に直接反映するモードは undo が世代を
  持つので何も失われない。原本を壊すのが怖くて `--copy` を選んだ人だけが成果を失う。

★ 2 冊の照合は同じ問題を既に解いていた（名前に依頼の条件を埋め、同じ条件の時だけ
  「前回の照合出力『…』を作り直しました」と言う）。単一ブックの経路に来ていなかっただけ
  ── `out_book_path` の docstring には、同じ関数が 2026-08-25 にも片配線の側だったと
  書いてある（フォルダ経路には関所が在り、単一ブック経路に無かった）。**3 度目**。

★ 直しは 2 つ（Namakoo 決裁 A＋C）:
  A 置き換える前に言う ── ただし**別の依頼で作った物を消す時だけ**。同じ依頼の作り直しで
    毎回鳴らすと、1 体目が言った「★ が毎回出るので読まなくなった」を繰り返す。
    言う材料は既に手元に在った（履歴の項目が前回の依頼文を持っている）。
  C `--out` を足す ── A の告知が「行動」に変わるのはこれが在る時だけ。
    ただし `.out` の名前は 118 の検体と原本反映の作業ファイルが依存しているので
    **既定は変えない**。行き先を選べる口だけを足す。

★ 行き先を決める式は `out_book_path(book)` が **4 箇所に書き写されて**いた。そこへ
  `--out` を配ると、配り忘れた経路だけが黙って `.out` に書く ── 決める場所を
  `run_output_path` 1 つに畳んでから配線した。
"""
import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _a(**kw):
    base = {"out": None, "copy": False, "task": ""}
    base.update(kw)
    return argparse.Namespace(**base)


# --- C: 行き先を決める場所は 1 つ ---------------------------------------------------

def test_the_destination_is_decided_in_one_place():
    """★★ 4 箇所に書き写された式が残っていないこと（配り忘れた経路が黙って .out に書く）。"""
    from _product_source import count_in_product
    assert count_in_product("out_book = run_output_path(a, book)") == 4, (
        "行き先を決める呼び出しが 4 経路そろっていない")
    assert count_in_product("out_book = out_book_path(book)") == 0, (
        "★ 古い式が残っている経路がある ── そこだけ --out が効かない（片配線）")


def test_without_the_flag_the_destination_is_unchanged(tmp_path):
    """★ 陰性対照: `--out` が無ければ今までどおり `<book>.out.xlsx`（契約は動かさない）。"""
    book = tmp_path / "売上.xlsx"
    assert ailine.run_output_path(_a(), book) == tmp_path / "売上.out.xlsx"


def test_a_bare_name_lands_next_to_the_book(tmp_path):
    """★ 名前だけ渡されたら**原本の隣**に置く（人が打つのは「集計.xlsx」であって path ではない）。"""
    book = tmp_path / "売上.xlsx"
    assert ailine.run_output_path(_a(out="集計.xlsx"), book) == tmp_path / "集計.xlsx"


def test_a_path_is_honoured(tmp_path):
    """★ 場所まで書かれていたらそこへ。"""
    book = tmp_path / "売上.xlsx"
    want = tmp_path / "sub" / "集計.xlsx"
    assert ailine.run_output_path(_a(out=str(want)), book) == want


def test_the_flag_only_makes_sense_with_copy():
    """★★ `--out` は `--copy` の時だけ（原本直接モードの `.out` は作業ファイル）。

    ★ 判定元は `--copy` そのもの ── `a.inplace` が立つのはこの検査より**後**で、
      そちらを見ると必ず False になり検査が素通りする（2026-09-17 に実測で踏んだ）。
    """
    from _product_source import count_in_product
    assert count_in_product('if getattr(a, "out", None) and not getattr(a, "copy", False):') == 1, (
        "★ 併用の検査が『--copy そのもの』を見ていない（inplace を見ると素通りする）")


# --- A: 別の依頼で作った物を消す時だけ言う -------------------------------------------

def test_it_says_what_it_is_about_to_replace(tmp_path, monkeypatch, capsys):
    """★ 事故そのもの: 別の依頼で作った `.out` を置き換える時は、書く前に言う。"""
    book = tmp_path / "b.xlsx"
    out = tmp_path / "b.out.xlsx"
    book.write_bytes(b"x")
    out.write_bytes(b"y")
    monkeypatch.setattr(ailine, "note_pre_existing_output", lambda _p: None)
    monkeypatch.setattr(ailine, "_file_digest", lambda _p: "SAME")
    monkeypatch.setattr(ailine, "read_history", lambda max_n=None: [
        {"out": str(out.resolve()), "out_sha": "SAME", "task": "仕入先ごとに金額を集計して"}])
    rc = ailine.refuse_if_output_is_someone_elses(book, "状態が未の行を抜き出して")
    said = capsys.readouterr().out
    assert rc is None, f"止めてしまった: {rc}"
    assert "置き換えます" in said, f"黙って置き換えている: {said!r}"
    assert "仕入先ごとに金額を集計して" in said, f"何を消すのか言っていない: {said!r}"
    assert "--out" in said, f"逃げ道を示していない: {said!r}"


def test_rebuilding_the_same_request_stays_quiet(tmp_path, monkeypatch, capsys):
    """★★ 陰性対照: 同じ依頼の作り直しでは黙る。

    ★ ここが無いと「毎回言う」でも上の試験が通る ── それは 1 体目の買い手が
      「★ が毎回出るので読まなくなった」と言った失敗をもう一度やる形。
    """
    book = tmp_path / "b.xlsx"
    out = tmp_path / "b.out.xlsx"
    book.write_bytes(b"x")
    out.write_bytes(b"y")
    monkeypatch.setattr(ailine, "note_pre_existing_output", lambda _p: None)
    monkeypatch.setattr(ailine, "_file_digest", lambda _p: "SAME")
    monkeypatch.setattr(ailine, "read_history", lambda max_n=None: [
        {"out": str(out.resolve()), "out_sha": "SAME", "task": "同じ依頼"}])
    rc = ailine.refuse_if_output_is_someone_elses(book, "同じ依頼")
    said = capsys.readouterr().out
    assert rc is None, f"止めてしまった: {rc}"
    assert "置き換えます" not in said, f"同じ依頼の作り直しで鳴った: {said!r}"


def test_an_edited_output_is_still_refused(tmp_path, monkeypatch):
    """★ 2026-08-25 の線を落としていないこと ── 人が手を入れた出力は消さずに断る。"""
    book = tmp_path / "b.xlsx"
    out = tmp_path / "b.out.xlsx"
    book.write_bytes(b"x")
    out.write_bytes(b"y")
    monkeypatch.setattr(ailine, "note_pre_existing_output", lambda _p: None)
    monkeypatch.setattr(ailine, "_file_digest", lambda _p: "CHANGED")
    monkeypatch.setattr(ailine, "read_history", lambda max_n=None: [
        {"out": str(out.resolve()), "out_sha": "SAME", "task": "前の依頼"}])
    assert ailine.refuse_if_output_is_someone_elses(book, "別の依頼") == 7
