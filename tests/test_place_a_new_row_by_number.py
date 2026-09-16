# まだ表に無い行を、行番号で置く ── 2026-08-31。
# Namakoo が提案した実演の 1 幕目（丸山工業を追加して値を入れていく）を、
# 本番の前に俺が先に通したところ **5 件中 5 件が落ちた**。
#
# ★★ ① 名前の切り出しが行番号をまたいで飲み込んでいた
#     「**8行目に**丸山工業の行を作って」
#       → 『8行目に丸山工業』という行が見つかりません
#     `_re_row_of`（「〜の行」）の区切りが空白と読点しか無く、助詞をまたいで拾っていた。
#     ★ そのとき task_names_a_row_number は正しく 8 を返していた ──
#       **行番号が分かっているのに、名前の切り出しがそれを無視していた。**
#
# ★★ ② 名前が表に無いことを、断りの理由にしていた
#     「これから置く」行なのだから、名前が表に無いのは**当たり前**。
#     依頼文が行番号を名指ししているなら、それが場所。
#
# ★ ①の直しは語彙ではなく**文法の線**（助詞は名前に含まれない）。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HEADERS = ["取引先", "項目", "件数", "単価", "金額"]
ROWS = [["丸和物流", "配送", 12, 4800, 57600],
        ["近江スチール", "鋼材", 5, 12000, 60000],
        ["ヤマノ食品", "食品", 28, 1500, 42000]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["請求"], "headers": {"請求": list(HEADERS)},
            "header_rows": {"請求": 1}, "path": str(p)}


# --- ① 名前に助詞を飲み込まない ------------------------------------------------------------

@pytest.mark.parametrize("task, want", [
    ("8行目に丸山工業の行を作って", "丸山工業"),
    ("丸和物流の行を削除して", "丸和物流"),
    ("5行目にヤマノ食品の行を追加して", "ヤマノ食品"),
])
def test_the_row_name_stops_at_a_particle(task, want):
    m = ailine._re_row_of.search(task)
    assert m and m.group(1) == want, m.group(1) if m else None


# --- ② これから置く行は、行番号で場所が決まる ----------------------------------------------

def test_a_new_name_with_a_row_number_resolves(meta):
    """★★ 実演の 1 幕目そのもの ── 名前が表に無くても、行番号が在れば置ける。"""
    at, note = ailine.resolve_row_anchor("8行目に丸山工業の行を作って", meta, "請求")
    assert at == 8, (at, note)
    assert "8行目" in note


def test_a_new_name_without_a_row_number_still_refuses(meta):
    """★ 黙りすぎていないこと: 場所の手掛かりが無ければ、今までどおり断る。"""
    at, note = ailine.resolve_row_anchor("丸山工業の行を作って", meta, "請求")
    assert at is None
    assert "見つかりません" in (note or ""), note


def test_an_existing_name_still_wins(meta):
    """★ 表に在る名前は今までどおり中身で解く（行番号に横取りされない）。"""
    at, note = ailine.resolve_row_anchor("丸和物流の行を削除して", meta, "請求")
    assert at == 2, (at, note)


def test_the_header_row_is_not_a_place(meta):
    """★ 見出し行を場所にしない。"""
    at, _note = ailine.resolve_row_anchor("1行目に丸山工業の行を作って", meta, "請求")
    assert at != 1


# --- ③ 実物で（実演の 1 幕目を通す）--------------------------------------------------------

@pytest.mark.local
def test_the_first_act_of_the_demo_runs(tmp_path):
    """★★ 行を作って、セルを 1 つ埋めるところまで（実演の入口）。"""
    import subprocess
    p = tmp_path / "d.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    env = {**__import__("os").environ, "PYTHONPATH": str(REPO / "src")}

    def _run(task):
        return subprocess.run(
            [sys.executable, "-m", "ailine", "run", str(p), task, "--sheet", "請求"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=900, cwd=str(REPO), env=env)

    r1 = _run("5行目に丸山工業の行を作って")
    assert "✓" in r1.stdout, r1.stdout[-1200:]
    r2 = _run("丸山工業の件数を20にして")
    assert "✓" in r2.stdout, r2.stdout[-1200:]
    v = openpyxl.load_workbook(p, data_only=True)["請求"]
    assert v.cell(5, 1).value == "丸山工業"
    assert v.cell(5, 3).value == 20


# ── ★★ 2026-09-16: 「印を付けて」が行追加に化け、台帳に架空の行が入った（買い手役 2 体目）──
#
# 実測（買い手の受注台帳・再現 2/2・**EXIT=0 で通っていた**）:
#     「納期が2026/09/30より後の行のチェック列に★を入れて」
#       → 解釈: 操作:行追加 挿入位置:6 入れる値:チェック=★
#       → 5 行目と 6 行目の間に、取引先も品目も金額も空の行が入った
#   ★ 買い手「印を付けてと頼んだだけです」「月末の受注台帳は、間違えたら請求が狂って、
#     取引先に謝りに行く紙です」。値段の回答は **0 円**で、理由はこの 1 件だった。
#
# ★ なぜ止まらなかったか: 「その op にしてよいか」を依頼文に問い返す器官（OP_META の
#   `requires_word`）は在るのに、宣言していたのは **PIVOT ただ 1 つ**だった。
# ★ 線は**実物で測って**引いた: repo の正当な行追加の検体に「〜列に／〜列を」は 1 件も無い。
#   行を足す依頼は「新しい行に入れる値」を言うのであって、書き込む先の列を名指ししない。

_META_ORDERS = {
    "sheets": ["受注一覧"],
    "headers": {"受注一覧": ["受注日", "取引先", "担当", "品目", "数量", "単価",
                             "金額", "納期", "状態", "チェック"]},
    "header_rows": {"受注一覧": 1},
}


@pytest.mark.parametrize("task", [
    "納期が2026/09/30より後の行のチェック列に★を入れて",          # ★ 事故そのもの
    "納期が2026/09/30より後の行だけ、チェック列を★に書き換えて",   # ★ 言い直しても同じだった
    "金額が50000以上の行のチェック列に★を入れて",                  # 比較語のある形
    "状態が未手配の行の備考列に「至急」と入れて",
    # ★ ここから下は「列」と書かない形 ── **比較語の枝だけ**が受け持つ。
    #   初版はこの形の検体が無く、比較語の枝を殺しても緑のままだった（在っても鳴らない規則）。
    "金額が50000以上の行に★を入れて",
    "数量が1000より多い行に「大口」と入れて",
])
def test_a_write_into_a_column_is_never_turned_into_a_new_row(task):
    """★ 列を名指しした依頼を**行追加にしない**。行が増えると、頼んでいない空行が台帳に入る。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": {"チェック": "★"}}, _META_ORDERS, task=task)
    assert not ok, f"列への書き込みを行追加として通した: {task}"
    assert "行を足す" in err, err
    assert "のように" in err, f"断るだけで、通る書き方を示していない: {err}"


@pytest.mark.parametrize("task", [
    "5行目に丸山工業の行を追加して",
    "1行目に丸山工業の行を作って",
    "3行目の下に新品を追加して",
    "北斗精機の行の下に、取引先「西村工業」の行を追加して、項目は事務机、件数は2、単価は15000にして",
])
def test_a_real_row_addition_still_passes_this_gate(task):
    """★ 陰性対照 ── repo が既に持っている**正当な行追加の言い方**を 1 件も落とさないこと。
       ★ この関所は「列を名指ししたか」だけを見る。値や位置の検査は従来どおり後ろで行う。"""
    from ailine_core import intent as intent_mismatch
    from ailine_core import compare_words
    refusal = intent_mismatch.refuse_row_add_that_is_really_a_write(
        task, comparison_in_task=compare_words.read(task).hit)
    assert refusal is None, f"正当な行追加を止めた: {task} → {refusal}"


def test_the_refusal_names_the_words_the_person_actually_wrote():
    """★ 断り文は、人が**自分の依頼文の中に見つけられる**語を名指しすること
       （初版は「…より後の行のチェック列に」を丸ごと拾って読めなかった）。

    ★ 2026-09-16 に助詞を落とした（「チェック列に」→「チェック列」）── 断り文が
       「『チェック列に』に値を入れる形」と二重になって日本語が壊れていたため。
       人が自分の依頼文の中に見つけられる、という趣旨は変わらない。
    """
    from ailine_core import intent as intent_mismatch
    for task, want in (
        ("納期が2026/09/30より後の行のチェック列に★を入れて", "チェック列"),
        ("納期が2026/09/30より後の行だけ、チェック列を★に書き換えて", "チェック列"),
        ("状態が未手配の行の備考列に「至急」と入れて", "備考列"),
    ):
        got = intent_mismatch.names_a_column_to_write_into(task)
        assert got == want, f"{task} → {got}"
