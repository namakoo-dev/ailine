# 利用者に見せる語彙と、モデルに見せる語彙が食い違っていないか ── 2026-08-30。
#
# ★★ Namakoo「語彙外の語については未対応か。改善できる可能性は残されてるか？」から
#   未測定だった op 系統を測った（bench/vocab_reach.py・48 件）。**到達率 79.2%**。
#   0/3 だったのは EXTRACT_COLUMNS（列抽出）と SET_WHERE（条件つき書換）。
#
# ★★ 原因は機械で確定できた ── **一覧が 2 つあって、片方だけ育っていた**:
#     登録簿 OP_META（`ailine ops` が利用者に見せる）  29 個
#     プロンプト OPS_DOC（モデルに見せる）             24 個
#   利用者には「できます」と言いながら、モデルには**その op の存在を教えていない**
#   ものが 5 つ（ADD_COLUMN / EXTRACT_COLUMNS / SET_CELL_VALUE / SET_WHERE / SWAP）。
#   ★ うち 3 つは機械の読み直しが後から拾っていたので動いていた。読み直しの無い
#     2 つ（列抽出・条件つき書換）だけが落ちていた ── **93 件の検体はそこを覆って
#     いなかったので、今まで見えなかった**。
#
# ★ さらに悪いことに、EXTRACT の説明は「一部の列だけを残す絞り込みは語彙に無い
#   （OUT_OF_VOCAB にする）」と**嘘を教えていた**（EXTRACT_COLUMNS は在る）。
#   0/3 のうち 2 件が OUT_OF_VOCAB だったのは、指示どおりに動いた結果。
#
# ★ これは「在っても鳴らない」の別の形 ── 二重に持っている一覧は必ずずれる。
#   ずれたら赤くする。

import re
import sys

import pytest
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _ops_in_prompt() -> set:
    return set(re.findall(r"^([A-Z_]{3,}):", ailine.OPS_DOC, re.M))


# ★★ 2026-08-30（教えたら悪化したので分けた）: 5 つ全部を OPS_DOC に足したところ、
#   到達率は 79.2% → 89.6% に上がったが、**完遂率が 98.9% → 82.8%・壊した数が 0 → 9**
#   になった（効果で測る 93 件）。原因:
#       4行目の数量を999にして      → 行数が変わった 4→5（行を挿した）
#       3行目の主材料を「東棟」にして → 1 行ずれた場所に書いた
#   モデルが SET_CELL_VALUE を**直接返す**ようになり、それまで機械の読み直しが
#   座標を解いていた経路を**素通り**した。★ 機械の解決のほうが正しかった。
#
# ★ だから「プロンプトに在ること」ではなく「**届く道があること**」を縛る。
#   道は 2 つ ── ①モデルに教える ②機械の読み直しが拾う。②で足りているものを
#   ①にも足すと、上のように素通りして悪くなる（実測）。
REACHED_BY_MACHINE_REREAD = {
    # op: 「なぜプロンプトに載せないか」＝載せたら悪くなったという実測
    "SET_CELL_VALUE": "載せると座標の解決を素通りして 4→5 行になった（実測 2026-08-30）",
    "SWAP": "軸（行か列か）は実表からしか決まらない。読み直しが解いている",
    "ADD_COLUMN": "位置も名前も機械が実表から決める。読み直しが解いている",
}


# ★★ 2026-08-30 の実測（3 通り・同じ条件）── **到達率と完遂率が逆相関した**:
#     教える op   到達率(48)   完遂率(93)   壊した
#     足さない     79.2%        98.9%        0     ← いまここ（採用）
#     2 つ足す     85.4%        96.8%        2
#     5 つ足す     89.6%        82.8%        9
#   ★ 語彙を広げると届きやすくなるが、**機械が持っていた確実な経路をモデルが奪う**。
#     実測: 「4行目の数量を999にして」で行数が 4→5、「3行目の主材料を…」で 1 行ずれた。
#     「氏名と所属の間に区分の列を追加して」は EXTRACT_COLUMNS に吸われた。
#   ★ **壊した数が 0 でなくなる取引は、この製品では成立しない** ── 戻した。
#
# ★ だから下の 2 つは **未解決として赤いまま残す**（xfail）。
#   隠すと「在っても鳴らない」を自分で作ることになる。直す道は分かっている:
#     ・機械の読み直しを EXTRACT_COLUMNS / SET_WHERE にも付ける（op を教えずに届かせる）
#   本番（2026-09-01）後に着手する。

UNREACHABLE_TODAY = {"EXTRACT_COLUMNS", "SET_WHERE"}


@pytest.mark.xfail(strict=True, reason=(
    "2026-08-30 実測: この 2 つはモデルにも教えておらず読み直しも無いので到達不能"
    "（48 件の測定で 0/3）。教えると完遂率が落ちるので、機械の読み直しを付ける方針。"))
def test_every_advertised_op_is_reachable():
    """★ `ailine ops` に載っている操作には、**届く道**があること。
       道は「モデルに教える」か「機械の読み直しが拾う」のどちらか。"""
    unreachable = sorted(set(ailine.OP_META) - _ops_in_prompt()
                          - set(REACHED_BY_MACHINE_REREAD))
    assert not unreachable, (
        "登録簿に在るのに、モデルにも教えず読み直しも無い op（到達不能）: "
        f"{unreachable}")


def test_the_unreachable_ones_are_exactly_the_ones_we_know_about():
    """★ 到達不能が**増えていない**ことを見張る（xfail で目をつぶる代わり）。"""
    unreachable = set(ailine.OP_META) - _ops_in_prompt() - set(REACHED_BY_MACHINE_REREAD)
    assert unreachable == UNREACHABLE_TODAY, (
        f"到達不能の顔ぶれが変わった: {sorted(unreachable)}")


def test_the_machine_rereads_actually_exist():
    """★ 上の免除表が**口約束になっていない**こと ── 読み直しが本当に在るかを見る。
       在っても鳴らない、の逆（無いのに在ると書く）を塞ぐ。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    for op in REACHED_BY_MACHINE_REREAD:
        assert f'plan = [{{"op": "{op}"' in src or f"'op': '{op}'" in src, (
            f"{op} は『読み直しで拾う』と書いてあるが、読み直しが見つからない")


def test_the_exemption_carries_its_reason():
    """★ 免除には**理由（＝実測）**を必ず添える（後から読む人が判断できるように）。"""
    for op, why in REACHED_BY_MACHINE_REREAD.items():
        assert len(why) > 10, op


def test_the_prompt_does_not_advertise_ops_that_do_not_exist():
    """★ 逆向き ── 存在しない op を教えると、走らせようがない計画が返る。"""
    extra = sorted(_ops_in_prompt() - set(ailine.OP_META))
    assert not extra, f"プロンプトに在るのに登録簿に無い op: {extra}"


def test_the_prompt_does_not_claim_a_real_op_is_out_of_vocabulary():
    """★★ 実測で出た形: EXTRACT の説明が『一部の列だけを残す絞り込みは語彙に無い』と
       書いていたが、EXTRACT_COLUMNS は**在る**。モデルは指示どおり OUT_OF_VOCAB を
       返していた ── 嘘を教えていたので 0/3。"""
    # ★ EXTRACT_COLUMNS 自体はまだ教えていない（上の xfail 参照）。
    #   ここで守るのは「**在るものを『無い』と教えない**」だけ ── 嘘は消した。
    i = ailine.OPS_DOC.index("EXTRACT:")
    seg = ailine.OPS_DOC[i:i + 700]
    assert "一部の列だけを残す絞り込み" not in seg, (
        "EXTRACT の説明が、実在する EXTRACT_COLUMNS を『語彙に無い』と教えている")
