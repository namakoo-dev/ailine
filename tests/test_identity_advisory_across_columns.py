# 等式の検算を、見出しの**名前**で対応づける番人（2026-09-02）。
#
# ★★ 自作 review が見つけた**致命**（敵対検証 2 レンズとも独立に再現）:
#   前日（f592ead）に入れた `if heads_b != _heads_a: return []` は、意図していた
#   「列の並べ替えで出る誤検知」だけでなく、**見出しが変わる操作すべて**
#   （列追加・列削除・見出しの変更）で `broken_identity_advisory` を丸ごと止めていた。
#   実測: 列を 1 本足すついでに直値の派生列（金額＝件数×単価）を壊しても ⚠ が出ない
#   ── **この関数が検出対象にしていた事故クラスを、最も普通の操作で握りつぶす**退行。
#
# ★ しかも同じ関数の docstring には
#     「op を問わず**1 箇所**で見る（入れ替えに限らず、入力を変える操作すべてに効く）」
#   と書いてあった ── **自己申告と実装が矛盾**していたのに、誰も突き合わせていなかった。
#
# ★ 直しは「降りる」ではなく**名前で対応づける**。位置でなく見出しで並べ直せば、
#   並べ替え・追加・削除を同じ手で扱える（前後の両方に在る列だけを比べる）。
#
# 契約:
#   ① 列を足しながら派生列を壊したら ⚠ が出る（元の設計目的）
#   ② 列を動かしただけでは鳴らない（前日に直した誤検知を戻さない）
#   ③ 列を足しただけでは鳴らない
#   ④ 列を消しながら壊しても ⚠ が出る
#   ⑤ 共通の列が 3 本に満たなければ黙る（等式が立たない）

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from _product_source import product_text  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

BASE = [["あ", 12, 4800, 57600], ["い", 5, 12000, 60000],
        ["う", 9, 7200, 64800], ["え", 3, 1000, 3000]]
HEADS = ["取引先", "件数", "単価", "金額"]


def _mk(path: Path, heads, rows) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(heads)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def _advisory(tmp_path, heads_after, rows_after):
    src = _mk(tmp_path / "src.xlsx", HEADS, BASE)
    out = _mk(tmp_path / "out.xlsx", heads_after, rows_after)
    return ailine.broken_identity_advisory(
        src, out, {"_target_sheet": "請求", "_header_row": 1})


def _broken_rows():
    rows = [r[:] for r in BASE]
    rows[0][3] = 99999                       # 金額 だけ壊す（件数×単価 と合わない）
    return rows


def test_a_break_while_adding_a_column_is_reported(tmp_path):
    """① 元の設計目的 ── 列を足すついでに派生列が壊れたら言う。

    ★ ここが review の指摘そのもの。前日の直しでは **[] を返して黙っていた**。
    """
    got = _advisory(tmp_path, HEADS + ["備考"], [r + [""] for r in _broken_rows()])
    assert got and "成り立たなくなりました" in got[0], got


def test_a_break_while_deleting_a_column_is_reported(tmp_path):
    """④ 逆向き ── 列を消しながら壊しても言う。"""
    rows = [[r[0], r[1], r[2], r[3]] for r in _broken_rows()]
    got = _advisory(tmp_path, ["件数", "単価", "金額"], [r[1:] for r in rows])
    assert got and "成り立たなくなりました" in got[0], got


def test_moving_a_column_is_still_silent(tmp_path):
    """② 前日に直した誤検知を戻さない ── 並べ替えただけでは鳴らない。

    ★ 「項目と件数を入れ替えて」で ⚠ が出た実測（Namakoo・実演の練習）への処置。
    """
    got = _advisory(tmp_path, ["取引先", "単価", "件数", "金額"],
                     [[r[0], r[2], r[1], r[3]] for r in BASE])
    assert got == [], got


def test_adding_a_column_alone_is_silent(tmp_path):
    """③ 足しただけでは鳴らない（鳴りすぎない側）。"""
    got = _advisory(tmp_path, HEADS + ["備考"], [r + [""] for r in BASE])
    assert got == [], got


def test_too_few_common_columns_is_silent(tmp_path):
    """⑤ 共通の列が 3 本に満たなければ黙る（等式が立たない）。"""
    got = _advisory(tmp_path, ["取引先", "件数"], [r[:2] for r in BASE])
    assert got == [], got


def test_the_docstring_promise_matches_the_code():
    """★ 自己申告と実装を突き合わせる ── review が突いたのはここだった。

    ★ 「op を問わず 1 箇所で見る」と書いてあるのに、見出しが変わると降りていた。
      **降りる分岐を戻したら赤くする。**
    """
    i = product_text().index("def broken_identity_advisory(")
    j = product_text().index(chr(10) + "def ", i + 10)
    body = product_text()[i:j]
    assert "op を問わず" in body, "宣言が消えている（消すなら実装も直すこと）"
    # ★ **コード行だけ**を見る ── コメントに同じ文字列が在っても赤くしない
    #   （初版はここで自分の説明文に引っかかった。番人が「コードか説明か」を
    #     区別していなかった ── 語で切ると必ずこうなる）。
    code = chr(10).join(ln for ln in body.splitlines()
                         if not ln.lstrip().startswith("#"))
    assert "if heads_b != _heads_a:" not in code, (
        "見出しが変わると降りる分岐が戻っている ── 列追加のついでの破損を握りつぶす")
