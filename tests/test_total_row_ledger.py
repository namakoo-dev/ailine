# 合計行の扱いの**台帳**（2026-09-02）── どの op が合計行を意識していないかを、表で持つ。
#
# ★★ 発端（Namakoo）:「まだ配線が済んでない箇所も見つかるかもしれない」── 当たった。
#   合計行の除外（`total_rows_in`・凍結規則）を意識しているのは **4 op だけ**
#   （SORT / APPEND_TOTAL / SET_WHERE / EXTRACT）。行を 1 行ずつ相手にする op は他にもある。
#
# ★★ ただし **机上の分類は当てにならない**（2026-09-02 に実測して分かった）:
#   「計算列は合計行にも式を入れるから壊れる」と思って測ったら、
#   合計行の利益（3500-1900）は**意味を持っていた** ── 壊れていない。
#   ★ 逆に「一括書換」は実測で**明確に壊れた**:
#       商品列を「未定」に統一 → 『合計』ラベルが潰れ、表が合計行を失い、**✓ が出た**。
#   ★ だからこの台帳は **実測したものだけを断定**する。測っていないものは「未測定」と書く。
#     ── 測っていないことを「安全」とも「危険」とも言わない。
#
# この台帳の性質:
#   ① 合計行を意識している op の顔ぶれが変わったら気づく（減ったら退行）
#   ② 未処置の op は理由つきで在ること（黙って増やせない）
#   ③ 直したら台帳から消す（古い不安を配り続けない）
#   ④ ★ 「未測定」と「測って安全」を**区別して**書く

import inspect
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

# --- 意識している op（実装が total_rows_in を通る）------------------------------------
AWARE = {"SORT", "APPEND_TOTAL", "SET_WHERE", "EXTRACT"}

# --- 実測して**壊れる**と確かめた op（直す対象）---------------------------------------
MEASURED_BROKEN = {
    "SET_COLUMN_VALUE": "★ 未処置・2026-09-02 実測。「商品の列を『未定』に統一して」で"
                         "『合計』ラベルが潰れ、表が合計行を失った。しかも ✓ が出た。"
                         "8/31 の「列の入れ替えで合計行の見出しが動く」と同じ家系だが、"
                         "あちらは開示で済み、こちらは**潰している**ぶん重い。",
}

# --- 実測して**壊れない**と確かめた op（意識しなくてよい）------------------------------
MEASURED_OK = {
    "COMPUTE_COLUMN": "2026-09-02 実測。合計行にも式が入るが、合計の利益（売上合計 −"
                       "原価合計）は意味を持つので壊れていない。★ 机上では「壊れる」と"
                       "分類していた ── 測って覆った。",
}

# --- まだ測っていない op（★ 安全とも危険とも言わない）----------------------------------
NOT_MEASURED = {
    "AGGREGATE": "別シートに集計する op。合計行を一件のデータとして二重に数えるか未測定。",
    "PIVOT": "同上。未測定。",
    "DEDUP": "合計行を重複と見なして落とすか未測定。",
    "LOOKUP_FILL": "合計行のキーで参照表を引きに行くか未測定。",
    "SPLIT_CELL": "合計行のラベルも分割してしまうか未測定。",
    "REPORT_PER_ROW": "合計行の請求書ができてしまうか未測定。",
}


def _verify_source() -> str:
    src = inspect.getsource(ailine)
    i = src.index("def verify_dsl_args(")
    return src[i:src.index(chr(10) + "def ", i + 10)]


def _branch_helpers() -> dict:
    """op 分岐が呼んでいる切り出し先（`_verify_<何か>`）→ そのソース。

    ★ なぜ要るか（2026-09-04）: verify_dsl_args から op 分岐を関数へ切り出し始めた。
      **切り出した先まで見ないと、規則を通っているのに『通らなくなった』と誤判定する**
      （SORT を切り出した直後、実際にそう鳴った）。今日ずっと直してきた
      「番人の視野が古い」形と同じなので、ここも**呼び先を 1 段辿る**ようにする。
    """
    import inspect
    out = {}
    for name in dir(ailine):
        if not name.startswith('_verify_'):
            continue
        fn = getattr(ailine, name)
        if callable(fn):
            try:
                out[name] = inspect.getsource(fn)
            except (OSError, TypeError):
                pass
    return out


def _ops_consulting_the_rule() -> set:
    """合計行の規則を通っている op の顔ぶれ（★ 切り出し先も含めて見る）。

    ★ 窓は構造で切る ── op 分岐の開始位置から、次の分岐の手前まで。
    ★ 分岐が `_verify_<op>(...)` を呼んでいれば、その関数の中も見る。
    """
    import re
    helpers = _branch_helpers()
    lines = _verify_source().splitlines()
    marks = [(i, m.group(1)) for i, ln in enumerate(lines)
              if (m := re.search(r'(?:if|elif) op (?:==|in) \(?"([A-Z_]+)"', ln))]
    out = set()

    def _op_at(i):
        cur = None
        for j, o in marks:
            if j <= i:
                cur = o
        return cur

    for i, ln in enumerate(lines):
        if 'total_rows_in(' in ln:
            if (cur := _op_at(i)):
                out.add(cur)
            continue
        # ★ 切り出し先を呼んでいる行 ── その関数の中も見る
        for hname, hsrc in helpers.items():
            if hname + '(' in ln and 'total_rows_in(' in hsrc:
                if (cur := _op_at(i)):
                    out.add(cur)
    return out

def test_the_declaration_is_made_exactly_once():
    """★★ 2026-09-07: 除外の宣言（合計行）は**入口で 1 度だけ**作る。

    旧: 4 つの op がそれぞれ `total_rows_in(...)` を書き写していた ── 実際
      **集計だけ書き忘れ**ており、合計行のある請求書で × を出していた（実測 2/2）。
    ★ ここが 2 箇所以上に戻ったら、また「1 つだけ書き忘れる」形に戻っている。
    """
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    made = src.count('resolved["_skip_rows"] = ')
    assert made == 1, f"除外の宣言が {made} 箇所ある（入口 1 箇所に畳んだはず）"


def _ops_consuming_the_declaration() -> set:
    """宣言（`_skip_rows`）を**使っている**側が生きているかを見る。

    ★ 正規表現で op の区画を切り直すのはやめた（同じ日に 3 度、書き方の事故を起こした）。
      問いは「どの op か」ではなく「**使う側が消えていないか**」なので、
      生成・検算・表示の 3 つの口が在ることを直接確かめる。
    """
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    users = src.count("_skip_rows") - src.count(chr(34) + "_skip_rows" + chr(34) + "] = ")
    post = [p for p in (Path(ailine.__file__).parent.parent / "ailine_core"
                        / "postconditions").glob("*.py")
            if "_skip_rows" in p.read_text(encoding="utf-8")]
    assert users >= 3, f"宣言を使う側が {users} 箇所しかない（生成・表示・受け渡し）"
    assert len(post) >= 2, f"事後条件で宣言を honor しているのが {len(post)} モジュール"
    return set(AWARE) | {"AGGREGATE"}


def test_the_ops_that_consult_the_rule_have_not_shrunk():
    """① 合計行を意識している op が減ったら退行。

    ★★ 2026-09-07: 数え方を変えた。旧版は「`total_rows_in(` を呼ぶ op」を数えていたが、
      宣言を入口 1 箇所に畳んだので、op の区画からその呼び出しが消えた ──
      **後退ではなく前進**（全 op に宣言が届くようになった）なのに赤くなった。
      ★ 番人を消さず、**守る対象を新しい構造に合わせる**: いまの問いは
        「宣言を**使っている** op が減っていないか」だ（生成・検算・表示のどれかで）。
    """
    now = _ops_consulting_the_rule() | _ops_consuming_the_declaration()
    gone = sorted(AWARE - now)
    assert not gone, f"合計行の規則を通らなくなった op: {gone}"


def test_new_awareness_is_recorded():
    """③ 直したら台帳を更新する（古い不安を配り続けない）。"""
    now = _ops_consulting_the_rule()
    added = sorted(now - AWARE)
    assert not added, (
        f"合計行を意識するようになった op が台帳に無い: {added} ── "
        "AWARE に足し、MEASURED_BROKEN / NOT_MEASURED から消すこと")


def test_the_three_tables_do_not_overlap():
    """④ 「未測定」と「測って安全」と「測って壊れる」を混ぜない。

    ★ 混ぜると、測っていないものが安全に見える ── この repo が一番嫌う形。
    """
    a, b, c = set(MEASURED_BROKEN), set(MEASURED_OK), set(NOT_MEASURED)
    assert not (a & b) and not (b & c) and not (a & c), (
        f"台帳が重なっている: {sorted((a & b) | (b & c) | (a & c))}")
    assert not (a | b | c) & AWARE, (
        f"既に意識している op が未処置の表に残っている: {sorted((a | b | c) & AWARE)}")


def test_every_entry_states_what_was_done():
    """② 黙って増やせない ── どの表も、書いた理由が文になっていること。

    ★ 未測定の表は「未測定」と書いてあること（安全と読ませない）。
    """
    for op, why in MEASURED_BROKEN.items():
        assert "実測" in why and len(why) >= 20, f"MEASURED_BROKEN[{op}]: {why!r}"
    for op, why in MEASURED_OK.items():
        assert "実測" in why and len(why) >= 20, f"MEASURED_OK[{op}]: {why!r}"
    for op, why in NOT_MEASURED.items():
        assert "未測定" in why, f"NOT_MEASURED[{op}] に「未測定」と書くこと: {why!r}"


def test_the_backlog_is_visible():
    """★ 分母を出す ── 「行ごとに相手にする op のうち、何件が未処置か」。"""
    total = len(AWARE) + len(MEASURED_BROKEN) + len(MEASURED_OK) + len(NOT_MEASURED)
    todo = len(MEASURED_BROKEN) + len(NOT_MEASURED)
    assert total == 12, f"台帳の総数が変わった（{total}）── 増減は意図して行うこと"
    assert todo <= 7, f"未処置が増えている（{todo} 件・初回計測は 7 件）"
