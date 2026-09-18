# -*- coding: utf-8 -*-
"""2 冊の照合も台帳に残る／履歴を組む場所は 1 つ（2026-09-18・盲検 3 体目 ⑨）。

★★ 起きたこと（製造業の購買担当・初見・盲検）:
  「**2 冊照合の履歴が `ailine history` に 1 件も出ません**（単独ブックの実行は全部出る）。
  **月次の証跡として使えない**」。

★★ またこの形だった ── 器官は在るが配線が無い。`_record_history` の docstring は
  2026-09-05 に**同じ事故**（CLARIFY が `_finish_run` を通らず台帳に一行も無かった）で
  こう書かれている:「記録する処理を 2 箇所に書くと片方だけ直る（この repo の系譜）。
  **1 本に畳んで呼び出し側に持たせない**」。にもかかわらず:

    ・csv 変換    手で dict を組む（out_sha あり）
    ・export-csv  手で dict を組む（out_sha **なし** ← 既に食い違っていた）
    ・照合        **呼び出しが無い**

  ★ out_sha は 2026-09-13 に買い手の致命で足した出力の指紋。**片方だけ直っていた**
    ── docstring の予言がそのまま起きている。
  ★ だから照合を 3 本目の写しにせず、`_record_side_command_history` 1 つへ畳んだ。

★ 畳む回に挙動は変えない ── export-csv は今も指紋を残さない（stamp_out=False）。
  その食い違いを直すかは**別の判断**として残す（頼まれていない範囲を勝手に広げない・
  2026-09-18 に桁区切りで 1 度やらかした）。

★ 断った回も残す ── 単一ブックの run は語彙外も台帳に残している。買い手が欲しいのは
  月次の証跡で、「何を頼んで通らなかったか」は成功と同じだけ要る。
"""
import ast
import contextlib
import io
import json
import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

A_ROWS = [["取引先", "金額"], ["甲社", 1000], ["乙社", 2000]]
B_ROWS = [["取引先", "金額"], ["甲社", 1000], ["乙社", 2000]]
#: ★ 断られる形は**列を増やして**作る ── 列が 2 本だけだと名指しが無くても型で一意に
#:   決まって通ってしまった（2026-09-18 に検体が甘くて赤になった・今日 3 度目）。
AMBIG_ROWS = [["取引先", "金額", "数量"], ["甲社", 1000, 3], ["乙社", 2000, 5]]


def _books(tmp_path, rows_a=None, rows_b=None):
    out = []
    for name, rows in (("a.xlsx", rows_a or A_ROWS), ("b.xlsx", rows_b or B_ROWS)):
        p = tmp_path / name
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in rows:
            ws.append(list(r))
        wb.save(p)
        out.append(p)
    return out


def _run(tmp_path, monkeypatch, task, rows_a=None, rows_b=None):
    """履歴を tmp へ逃がして照合を 1 回走らせ、書かれた行を返す。"""
    hist = tmp_path / "history.jsonl"
    monkeypatch.setattr(ailine, "HISTORY_FILE", hist)
    a, b = _books(tmp_path, rows_a, rows_b)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = ailine.main(["run", str(a), str(b), task])
    rows = [json.loads(ln) for ln in hist.read_text(encoding="utf-8").splitlines() if ln.strip()] \
        if hist.exists() else []
    return rc, rows, buf.getvalue()


def test_a_match_that_worked_is_recorded(tmp_path, monkeypatch):
    """★ 事故そのもの: 通った照合が台帳に残ること。"""
    rc, rows, out = _run(tmp_path, monkeypatch, "取引先をキーにして金額を突き合わせて")
    assert rc == 0, out
    assert rows, f"照合が台帳に 1 行も残っていない（月次の証跡にならない）:\n{out}"
    e = rows[-1]
    assert e["ok"] is True and e["path"] == "match", e
    assert e["task"] == "取引先をキーにして金額を突き合わせて", e
    assert e["out"] and "照合" in Path(e["out"]).name, e


def test_a_match_that_was_refused_is_recorded_too(tmp_path, monkeypatch):
    """★★ 断った回も残す ── 「何を頼んで通らなかったか」は証跡として成功と同じだけ要る。

    ★ 単一ブックの run は語彙外を残している。片方だけ残すのは、また片配線。
    """
    rc, rows, out = _run(tmp_path, monkeypatch, "よしなに突き合わせて",
                         rows_a=AMBIG_ROWS, rows_b=AMBIG_ROWS)
    assert rc == 3, out
    assert rows, f"断った照合が台帳に残っていない:\n{out}"
    e = rows[-1]
    assert e["ok"] is False and e["path"] == "match", e
    assert e["failure_kind"], "なぜ通らなかったかが残っていない"


def _calls(name: str) -> list:
    """製品コードの中で `name(...)` を**実際に呼んでいる**場所（AST）。

    ★★ 文字列で数えると、docstring やコメントに書いた**コード片まで数える** ──
      2026-09-18 にこの検体自身がそれで赤くなった（`stamp_out=False` が 3 件）。
      朝に `compare_words.read(` で踏んだのと同型で今日 4 度目。綴りでなく意味を数える。
    ★★ そして**読む場所も決め打ちしない** ── 初版は src/ailine/__init__.py を直に開いて
      いて、tests/test_guard_ledger.py に捕まった（「本体を場所で決め打ちする番人が増えた」）。
      2026-09-03 に事後条件を ailine_core/ へ移した時、同じ形で番人 7 件が同時に空振りして
      いる ── **番人自身の片配線**。`_product_source.product_files()` を通れば次にどこへ
      分割しても空振りしない。
    """
    from _product_source import product_files
    out = []
    for f in product_files():
        tree = ast.parse(f.read_text(encoding="utf-8"))
        out += [n for n in ast.walk(tree)
                if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == name]
    return out


def test_the_history_entry_is_built_in_one_place():
    """★★ 手で dict を組む場所は 1 つ ── 2 箇所に書くと片方だけ直る（実際そうなっていた）。"""
    raw = [c for c in _calls("append_history")
           if c.args and isinstance(c.args[0], ast.Dict)]
    assert len(raw) == 1, (
        f"★ 履歴の dict を手で組む場所が {len(raw)} 箇所ある"
        "（_record_side_command_history へ畳むこと）")


def test_the_side_routes_all_go_through_the_helper():
    """★ DSL を通らない経路が全部、同じ器を通ること。

    ★ 呼び手は 4 ── csv 変換 / export-csv / 照合の**成功** / 照合の**断り**。
      照合を 2 箇所に配線したのは、単一ブックの run が成功も語彙外も残しているのと
      揃えたから（片方だけ残すのは、また片配線）。
    """
    n = len(_calls("_record_side_command_history"))
    assert n == 4, (
        f"★ 呼び手が {n} 本（期待 4: csv / export-csv / 照合の成功 / 照合の断り）── "
        "どれかが畳まれていないか、配線が増減した")


def test_every_side_route_stamps_its_output():
    """★★ DSL を通らない経路は**全部**、出力の指紋（out_sha）を残すこと。

    ★ 2026-09-18（Namakoo 決裁 A）: export-csv だけ残していなかった ── 変換の側には
      2026-09-13 に買い手の致命で足したのに、**片方だけ直っていた**。揃えた。
    ★★ 到達の記録: この指紋には**まだ読み手が居ない**。`_csv_output_edited_since` は
      path=="csv" だけを見ており、書き出し系の関所は「在れば --overwrite を要求する」
      だけで指紋を見ない。**材料が揃っただけで、効いてはいない**。
      ★ 読み手を配線する日（別の判断）に、この検体が道しるべになる。
        「指紋が在るから守られている」と読ませないために、ここに書いておく。
    """
    off = [c for c in _calls("_record_side_command_history")
           if any(k.arg == "stamp_out" and getattr(k.value, "value", None) is False
                  for k in c.keywords)]
    assert not off, (
        f"指紋を残さない経路が {len(off)} 本ある ── 揃えた線が崩れている"
        "（意図した変更なら、この検体も一緒に直すこと）")
