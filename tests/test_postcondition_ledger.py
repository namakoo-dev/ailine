# 事後条件の**台帳**（2026-09-02）── どの op が何を確かめていないかを、表で持つ。
#
# ★★ 発端（Namakoo）:「うまく割れたら片配線のチェックをしたほうがいいな。
#   まだ配線が済んでない箇所も見つかるかもしれない」── その読みが当たった。
#
#   数えたら **28 op のうち 8 op が「他は 1 セルも変わらず」を見ていなかった**
#   （適用前のファイルを引数に受けてすらいない）。
#   ★ これは 2026-08-30 に計算列で見つけたのと**同じ形**が 8 箇所残っていたということ:
#     「1セル書換・行追加・行削除・並べ替え・入れ替えには在るのに、計算列だけ無かった
#       ── また片配線」。その時は 1 op だけ直して、**数えなかった**。
#
# ★ 「数えなかった」が今回の教訓そのもの。**穴を 1 つ塞ぐたびに、同じ形の穴を数える。**
#
# この台帳の性質:
#   ① 全 op が事後条件を持つ（持たないなら**理由つきで**ここに書く）
#   ② 全 op が「他は変わっていない」を見る（見ないなら**理由つきで**ここに書く）
#   ③ 免除は**増やせるが、黙っては増やせない** ── 理由の無い免除は赤くなる
#   ④ 免除が減ったら（＝直したら）台帳も減らす。★ 直したのに残っていると、
#      「まだ穴が在る」と嘘をつく台帳になる
#
# ★ 台帳は「決めた時に書く」方式にしない ── **機械が掃き出したものと突き合わせる**。
#   人が書き忘れたら、その分だけ静かに守られなくなる。

import inspect
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

# --- 免除の台帳 ------------------------------------------------------------------
#
# ★ 書き方: op → なぜ見られないのか（**技術的な理由**を書く。「まだ」は理由ではない）

NO_POSTCONDITION = {
    "CHART": "グラフは読み戻して形を確かめる手段が無い（openpyxl では図形の中身を"
              "取り出せない）。見た目の判定は人がやる、と README に明記している。",
}

# ★★ ここが 2026-09-02 に**数えて出てきた 8 件**。直すたびに 1 行ずつ消していく。
#   ★ 「他は 1 セルも変わらず」を見ていない ＝ 頼んでいない場所への書き込みを見逃す。
#     番人の感度を測る治具（bench/guard_sensitivity.py）が 2026-08-30 に
#     計算列でこの穴を見つけた ── **同じ形が 8 箇所残っていた**。
NO_UNCHANGED_CHECK = {
    "AGGREGATE": "★ 未処置。別シートを作る op。元シートが無傷であることを見ていない。",
    "APPEND_TOTAL": "★ 未処置。末尾に合計行を足す op。他の行が無傷であることを見ていない。",
    "DRAW_BORDERS": "★ 未処置。書式だけを変えるはずの op。**値が 1 つも変わっていない**"
                     "ことを見ていない（書式 op が値を壊したら、それは重い事故）。",
    "LOOKUP_FILL": "★ 未処置。1 列を埋める op。他の列が無傷であることを見ていない。",
    "MERGE": "★ 未処置。セルを結合する op。結合で消える値を見ていない。",
    "NUMBER_FORMAT": "★ 未処置。書式だけを変えるはずの op。値の不変を見ていない。",
    "PIVOT": "★ 未処置。別シートを作る op。元シートが無傷であることを見ていない。",
    "SET_COLUMN_VALUE": "★ 未処置。1 列を書き換える op。他の列が無傷であることを見ていない。",
}


def _body(fn) -> str:
    try:
        return inspect.getsource(fn)
    except (OSError, TypeError):
        return ""


def _sees_before(fn) -> bool:
    """適用前のファイルと突き合わせているか（引数に受けているか）。

    ★ 「他は 1 セルも変わらず」は**適用前が要る** ── 出力だけを見ても
      「頼んでいない場所が変わったか」は原理的に分からない。
    """
    try:
        return "source_book" in inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return False


def test_every_op_has_a_postcondition_or_a_written_reason():
    """① 事後条件が無い op は、理由つきで台帳に在ること。"""
    missing = sorted(set(ailine.OP_LABELS) - set(ailine.POSTCONDITIONS))
    unlisted = [op for op in missing if op not in NO_POSTCONDITION]
    assert not unlisted, (
        f"事後条件が無いのに台帳に理由が無い op: {unlisted} ── "
        "事後条件を書くか、なぜ書けないかを NO_POSTCONDITION に書くこと")


def test_every_op_compares_against_the_before_file_or_is_listed():
    """② 「他は変わっていない」を見ない op は、理由つきで台帳に在ること。

    ★ ここが**片配線の検出器**そのもの。1 op を直したときに、
      同じ形の穴が他に何個あるかを毎回数える。
    """
    lack = sorted(op for op, fn in ailine.POSTCONDITIONS.items() if not _sees_before(fn))
    unlisted = [op for op in lack if op not in NO_UNCHANGED_CHECK]
    assert not unlisted, (
        f"適用前と突き合わせていないのに台帳に無い op: {unlisted} ── "
        "source_book を受けて比べるか、なぜ比べられないかを NO_UNCHANGED_CHECK に書くこと")


def test_the_ledger_does_not_keep_stale_entries():
    """④ 直したのに台帳に残っていたら赤くする。

    ★ これが無いと台帳は**古い不安を配り続ける**（「まだ穴が在る」という嘘）。
      在庫の数が減ったことを、機械が確かめる側。
    """
    lack = {op for op, fn in ailine.POSTCONDITIONS.items() if not _sees_before(fn)}
    stale = sorted(set(NO_UNCHANGED_CHECK) - lack)
    assert not stale, (
        f"もう比べている op が台帳に残っている: {stale} ── 直したら台帳からも消すこと")
    stale2 = sorted(set(NO_POSTCONDITION) & set(ailine.POSTCONDITIONS))
    assert not stale2, f"事後条件が付いた op が台帳に残っている: {stale2}"


def test_every_exemption_states_a_technical_reason():
    """③ 免除は増やせるが、**黙っては増やせない**。

    ★ 「まだ」「あとで」は理由ではない ── それは未処置の告白であって、
      **未処置と分かる形で書く**なら良い（★ 未処置。で始める約束にする）。
    """
    for table, label in ((NO_POSTCONDITION, "NO_POSTCONDITION"),
                          (NO_UNCHANGED_CHECK, "NO_UNCHANGED_CHECK")):
        for op, why in table.items():
            assert len(why) >= 20, f"{label}[{op}] の理由が短すぎる: {why!r}"
            assert "。" in why, f"{label}[{op}] の理由が文になっていない"


def test_the_count_of_unwired_ops_is_visible():
    """★ 分母を出す ── 「8 件中いくつ直したか」が常に見える形にする。

    ★ この数が減ることが、片配線を潰した証拠になる。増えたら気づく。
    """
    lack = {op for op, fn in ailine.POSTCONDITIONS.items() if not _sees_before(fn)}
    assert len(lack) == len(NO_UNCHANGED_CHECK), (
        f"未配線 {len(lack)} 件 / 台帳 {len(NO_UNCHANGED_CHECK)} 件 ── 数が合わない")
    # ★ 2026-09-02 の初回計測: 28 op 中 8 op が未配線だった。
    #   ここを下げていくのが仕事。**上がったら赤くする。**
    assert len(lack) <= 8, f"未配線が増えている（{len(lack)} 件・初回計測は 8 件）"
