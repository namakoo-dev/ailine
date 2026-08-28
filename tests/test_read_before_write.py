# 書く前に「こう読みました」を見せる ── 2026-08-28。
#
# ★★ なぜ要るか（Namakoo と一致した認識）:
#   「表記ゆれは矯正しても他の箇所で表記ゆれを起こす可能性を否定できない」
#   実際、この 2 日で見つかった壊れの大半が言い回しの揺れ由来だった:
#     ・「◎を付けて」は通るが「◎を入れて」は通らない
#     ・「7行目の担当を『佐藤』に」で**担当列が全行書き換わって ✓ が出た**
#     ・「7 行F列に『佐藤』を追加」が 3/3 とも ADD_ROW（行の追加）に化けた
#   ★ 事後条件は**宣言と実体**しか照らせない。**依頼と宣言**を照らせるのは人だけ。
#     だから揺れを 1 つずつ矯正するのではなく、**人が照らす機会を書く前に渡す**。
#
# 契約:
#   ① `--dry` は 1 バイトも書かない（画面がそう言い切るので、機械で縛る）
#   ② `--op` で人が操作を固定できる（当てる段を飛ばし、第二段に args だけ埋めさせる）
#   ③ 固定した回は**黙って別の op に読み直さない**（画面に出した読みと実行が食い違わない）
#   ④ ただし「1 行を指す依頼」を列ぜんぶ書き換える op で走らせるのは断る
#   ⑤ 画面は本体が出した行をそのまま見せる（画面で言い換えない）

import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _book(tmp_path, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(["取引先", "件数", "担当"])
    for row in [["丸和物流", 10, "田中"], ["ヤマノ食品", 20, "鈴木"], ["北斗精機", 30, "田中"]]:
        ws.append(row)
    wb.save(p)
    return p


# --- ② CLI の入口 ---------------------------------------------------------------------

def test_the_run_command_takes_a_fixed_op():
    """★ 画面の「選び直す」はこの旗を通る ── 旗が消えたら画面の道も消える。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert 'r.add_argument("--op"' in src, "--op が無い"
    i = src.index('forced_op = getattr(a, "op", None)')
    seg = src[i:i + 1400]
    assert "translate_task_fixed_op" in seg, "固定 op が第二段へ配線されていない"
    assert "OP_SCHEMA" in seg, "実在しない op 名を弾いていない"


def test_an_unknown_op_is_refused(tmp_path):
    p = _book(tmp_path)
    r = subprocess.run([sys.executable, "-m", "ailine", "run", str(p), "担当を佐藤に",
                         "--op", "NOSUCH", "--dry"],
                        capture_output=True, text=True, encoding="utf-8", errors="replace",
                        timeout=120, cwd=str(REPO))
    assert r.returncode == 3, r.stdout[-400:]
    assert "そんな操作はありません" in r.stdout, r.stdout[-400:]


# --- ①「まだ 1 バイトも変わっていません」を機械で縛る -----------------------------------

def test_a_dry_read_does_not_touch_the_file(tmp_path):
    """★★ 画面がそう言い切る文言なので、言い切りの裏を機械で取る。
       ここが破れると、確認のつもりで押した人のファイルが変わる。"""
    p = _book(tmp_path)
    before = p.read_bytes()
    r = subprocess.run([sys.executable, "-m", "ailine", "run", str(p),
                         "担当を全部「佐藤」にして", "--dry"],
                        capture_output=True, text=True, encoding="utf-8", errors="replace",
                        timeout=300, cwd=str(REPO))
    assert r.returncode in (0, 3), r.stdout[-400:]
    assert p.read_bytes() == before, "--dry なのにファイルが変わった"
    assert not list(p.parent.glob("*.out.xlsx")), "--dry なのに出力ができた"


# --- ③④ 人の選択を、黙って書き換えない -------------------------------------------------

def test_a_forced_op_skips_every_reread():
    """★★ 画面に「こう読みました」と出したあとで別の op に化けたら、その表示は嘘になる。
       読み直しの層は**印 1 つ**で丸ごと止める（層が増えても止まり続ける）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert '_reread_done = bool(getattr(a, "_forced_op", None))' in src, \
        "固定 op のときに読み直しを止めていない"


def test_pointing_at_one_row_with_a_column_wide_op_is_refused():
    """★ ④: 人が『一括書換』を選んでも、依頼が 1 行を指しているなら断る
       （画面に出した読みと結果が食い違う唯一の道を塞ぐ）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index('forced_op = getattr(a, "op", None)')
    seg = src[i:i + 1600]
    assert "plan_writes_beyond_one_cell" in seg and "task_points_at_one_row" in seg, seg[:400]
    assert "1セル書換" in seg, "選び直す先を名指ししていない"


# --- ⑤ 画面（薄い殻のまま）-------------------------------------------------------------

def test_the_screen_asks_before_it_writes():
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    assert 'id="readpanel"' in html and "こう読みました" in html
    # 「下書きで実行」は、いきなり書かずにまず読む
    i = html.index('$("#runcopy").onclick')
    assert "readFirst(" in html[i:i + 200], html[i:i + 200]
    # 実行に進むのは、人が押したときだけ
    j = html.index('$("#readgo").onclick')
    assert '"/api/run"' in html[j:j + 300], html[j:j + 300]


def test_the_screen_does_not_paraphrase_the_reading():
    """★ 画面が読みを言い換えた瞬間に嘘が入る ── 本体が出した行をそのまま映す。"""
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    i = html.index("async function readFirst(")
    seg = html[i:i + 1800]
    assert 'startsWith("解釈:")' in seg, seg[:400]
    assert "el.textContent = ln" in seg, "本体の行をそのまま出していない"


def test_a_refusal_disables_the_run_button():
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    i = html.index("async function readFirst(")
    seg = html[i:i + 1800]
    assert '$("#readgo").disabled = refused' in seg, seg[-400:]


def test_the_server_has_a_read_only_endpoint():
    src = (REPO / "gui" / "server.py").read_text(encoding="utf-8")
    i = src.index('if u.path == "/api/read":')
    seg = src[i:i + 900]
    assert '"--dry"' in seg, "読むだけの入口が --dry を渡していない"
    assert "_op_args(req)" in seg and "_sheet_args(req)" in seg, seg[:400]
    # ★ 恒真殺し: 読むだけの入口が、下書きを作る側の道に紛れ込んでいないこと
    assert "_DRAFTS[" not in seg, "読むだけの入口が下書きを作っている"


def test_the_op_picker_is_shared_with_the_alias_form():
    """★ 選択肢は本体の登録簿から取る（op が増えた日に、画面も黙って増える）。"""
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    assert 'fillOpPicker($("#readop")' in html
    assert '"/api/oplist"' in html
