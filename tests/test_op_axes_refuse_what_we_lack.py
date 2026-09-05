"""op の軸の上で**まだ**扱えないものは、そう言って断ること（2026-09-05）。

★★ 出所: 断りの導線を 15 件で測ったら、提案の 3/9 が**意図とずれていた**。

    「単価の平均値を一番下に追加して」→ もしかして: 合計追加？
    （『平均値・一番下・追加』などの部分はこの操作に反映されません）

  ★ 道具は**自分でずれを分かっている**のに、提案として先頭に出していた。
★ Namakoo:「候補（提案の出し方を変える等）はどれも逃げで、問題を解いていない」──
  正しい応答は「**平均はまだ扱えません**」と言うこと。

★★ 構造の穴: 知識は語彙表に**散文**で在った（「合計(SUM)専用。平均・最大・最小は語彙に無い」）。
  これは LLM への**指示**であって保証ではない ── 実際、判定器はこの文を読んでいるのに
  APPEND_TOTAL を出した。機械が読む宣言（ailine_core/op_axes.py）にした。

★ 2026-08-22 に却下された「道具が持たない全機能の名簿」（**開集合**・増築しても収束しない）
  とは別物。ここは**その op の軸の上の兄弟**という**閉集合**で、軸は op の意味が決める。

★★ 実装前に紙の上で測って形を決めた（今日 A′ でやった作法）:
    単純な部分文字列一致  → **5/6 で誤爆**（「平均単価の合計」を断る等）
    実表に聞く版          → 9/10
    列名が依頼文に在るかで締めた版 → 11/12
    ★ ambiguous を足した版（Namakoo の指示）→ **13/13**

★ いちばんこわいのは「**持っている機能を断る**」側 ── 下の 2 群を**対で**縛る。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402
from ailine_core.op_axes import (  # noqa: E402
    AXES, judge_axis, render_axis_ambiguity, render_axis_refusal)

#: 「平均単価」「余り」など、**持たない語を含む列名**を持つ表（誤爆の温床）
HEADERS = ["品名", "平均単価", "最大在庫数", "件数", "余り", "金額"]


# --- ① まだ無いものは断る --------------------------------------------------

@pytest.mark.parametrize("op, task, want", [
    ("APPEND_TOTAL", "単価の平均値を一番下に追加して", "平均"),
    ("APPEND_TOTAL", "単価の最大値を一番下に追加して", "最大"),
    ("APPEND_TOTAL", "中央値を一番下に追加して", "中央値"),
    ("APPEND_TOTAL", "単価の合計と平均を出して", "平均"),
    ("COMPUTE_COLUMN", "数量の累乗の列を作って", "累乗"),
    ("DEDUP", "重複を全部消して", "重複を全部消す"),
])
def test_something_we_do_not_have_yet_is_refused(op, task, want):
    verdict, detail = judge_axis(op, task, HEADERS)
    assert verdict == "lacks", (task, verdict, detail)
    assert detail[0] == want, detail
    said = "\n".join(render_axis_refusal(op, detail))
    assert want in said and "まだ扱えません" in said, said


# --- ② 持っているものは通す（★ こちらの誤爆が一番こわい）--------------------

@pytest.mark.parametrize("op, task", [
    ("APPEND_TOTAL", "金額の合計を一番下に追加して"),
    ("APPEND_TOTAL", "税込み合計を出して"),
    ("APPEND_TOTAL", "平均単価の合計を一番下に追加して"),   # ★ 列名に「平均」
    ("APPEND_TOTAL", "最大在庫数の合計を出して"),           # ★ 列名に「最大」
    ("APPEND_TOTAL", "件数と金額の合計を出して"),           # ★ 列名が「件数」
    ("COMPUTE_COLUMN", "数量と単価をかけた金額の列を作って"),
    ("COMPUTE_COLUMN", "余りの列と数量をかけた列を作って"),  # ★ 列名が「余り」
    ("DEDUP", "品名が同じ行を重複として除いて"),
    ("DEDUP", "全部消さずに重複だけ除いて"),                # ★ 否定形
    ("SORT", "金額の大きい順に並べ替えて"),                 # ★ 軸に lacks が無い op
])
def test_something_we_do_have_is_not_refused(op, task):
    """★ 単純な部分文字列一致だと 5/6 で誤爆した群 ── 実表に聞いて避ける。"""
    verdict, detail = judge_axis(op, task, HEADERS)
    assert verdict == "ok", (task, verdict, detail)


# --- ③ 決められない時は「判断しかねる」と言う（★ Namakoo の指示）------------

def test_a_word_that_is_both_a_column_and_an_operation_is_declared_ambiguous():
    """★ 黙って通すより正直 ── 列名と操作語が同居すると機械には決められない。

    ★ ナギは当初「見逃す（＝いまと同じ挙動なので悪化しない）」で妥協しようとしたが、
      Namakoo の「循環する語は警告を出して言い直し推奨」で解が完成した。
      実装の複雑さは同じだった。
    """
    verdict, detail = judge_axis("APPEND_TOTAL", "平均単価の平均を出して", HEADERS)
    assert verdict == "ambiguous", (verdict, detail)
    said = "\n".join(render_axis_ambiguity("APPEND_TOTAL", detail))
    assert "判断しかねます" in said and "平均単価" in said, said
    assert "言い方に直して" in said, said


# --- ④ 宣言そのものの形 ----------------------------------------------------

def test_every_axis_declares_what_it_has_and_why_it_is_pending():
    """★ 「持たない」でなく「**まだ無い**」── 発火条件つきで残すこと（Namakoo）。"""
    for op, axis in AXES.items():
        assert axis.name and axis.has, op
        assert axis.lacks, f"{op}: lacks が空なら軸を宣言する意味が無い"
        assert axis.note and "保留" in axis.note, f"{op}: 保留の理由と発火条件が無い"


def test_axes_only_name_real_ops():
    for op in AXES:
        assert op in ailine.OP_SCHEMA, f"知らない op: {op}"


def test_the_judgement_needs_the_real_table():
    """★ 実表を渡さないと誤爆する ── そのことを試験で示しておく。

    ★ headers 無しでは「平均単価の合計」を断ってしまう（実装前の測定で 5/6 誤爆）。
      呼び出し側が必ず実表を渡すことは下の試験が縛る。
    """
    assert judge_axis("APPEND_TOTAL", "平均単価の合計を出して", [])[0] == "lacks"
    assert judge_axis("APPEND_TOTAL", "平均単価の合計を出して", HEADERS)[0] == "ok"


def test_callers_pass_the_headers():
    """★ 呼び出し側が実表の見出しを渡していること。"""
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    n = src.count("judge_axis(")
    assert n >= 2, "配線が足りない"
    for chunk in src.split("judge_axis(")[1:]:
        head = chunk[:120]
        assert "_ax_head" in head or "_ax_headers" in head, f"実表を渡していない: {head[:60]}"


@pytest.mark.local
def test_the_refusal_shows_a_wording_that_actually_works(tmp_path):
    """★ 実機 ── 断りに出る例が**そのまま通る**こと（今日直した examples と繋がる）。"""
    import openpyxl, os, re, subprocess
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "T"
    ws.append(["品名", "平均単価", "金額"]); ws.append(["机", 100, 12000])
    src = tmp_path / "in.xlsx"; wb.save(src)
    repo = Path(ailine.__file__).resolve().parents[2]
    env = {**os.environ, "PYTHONPATH": str(repo / "src")}

    def _run(task):
        return subprocess.run([sys.executable, "-m", "ailine", "run", str(src), task,
                               "--copy", "--sheet", "T", "--timeout", "150"],
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", cwd=str(repo), env=env)

    got = _run("単価の平均値を一番下に追加して")
    assert got.returncode == 3, got.stdout[-400:]
    assert "まだ扱えません" in got.stdout, got.stdout[-400:]
    shown = re.findall(r"「([^」]{6,})」", got.stdout)
    assert shown, f"言い直しの例が出ていない: {got.stdout[-400:]}"
    again = _run(shown[0])
    assert again.returncode == 0, f"示した例が通らない: 「{shown[0]}」\n{again.stdout[-400:]}"
