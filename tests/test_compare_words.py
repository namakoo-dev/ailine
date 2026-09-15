# -*- coding: utf-8 -*-
# 辞書を「正確に引く」機構の番人（2026-09-15・設計と凍結した予測は docs/DESIGN-20260915-辞書を正確に引く.md）。
#
# ★ なぜ在るか: 言い回し 120 件の盲検で辞書に口語を足したあと、Namakoo「辞書登録後は正確に
#   引いてこれる仕組みも必要だ」。引き方を実測したら欠陥が 4 つ、どれも**辞書を増やすほど悪化**:
#     「20時間を超えない人」→ gt（逆）／「3000以上5000未満」→ gte（未満を黙って捨てる）／
#     「10を切っていない品番」→ None → LLM の不等号がそのまま通る
#
# 番人は 5 つ（辞書が育っても壊れない形）:
#   1. 1 語 1 検体 ── 全レコードの陽性を機械で回す（検体の無い語は入らない）
#   2. 到達できない語は入らない ── 自分の検体で**その語が**勝つ
#   3. 語彙の共食いの凍結 ── 402 句の読み（bench/compare_words_freeze.json）が 1 件も動かない
#   4. 出所の無い語は入らない
#   5. ★ 陰性対照と活用の表（2026-09-15 の盲検レビューが捕まえた 2 つの穴の番人）
#      ── 凍結 402 句は比較を含む句が 5% しか無く、否定検出と COUNTER ガードを丸ごと消しても
#         1 件も動かなかった（在っても鳴らない番人）。だから**鳴る検体を名指しで持つ**
import json
import os
import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import compare_words as C  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _product_source import count_in_product  # noqa: E402 ── ★ 本体を場所で決め打ちしない

FREEZE = REPO / "bench" / "compare_words_freeze.json"
REGEN = os.environ.get("AILINE_REGEN_COMPARE_FREEZE") == "1"


def _shown(msg: str) -> list:
    """断り文の**可変部**（『…』『…』）だけを取り出す ── 固定文の語（「以上」「範囲」）で恒真にしない。"""
    m = re.search(r"（『(.+?)』）", msg)
    return m.group(1).split("』『") if m else []


# ── 辞書そのものの番人 ────────────────────────────────────────────────────────

def test_every_word_has_a_specimen_and_wins_it():
    """1・2: 全レコードに陽性 1 件が在り、その検体で**その語が**勝つ（他の語に覆われていない）。"""
    bad = []
    for w in C.WORDS:
        if not w.example:
            bad.append((w.text, "検体が無い"))
            continue
        r = C.read(w.example)
        if r.cmp != w.cmp or r.word != w.text:
            bad.append((w.text, f"自分の検体で勝てない: {r}"))
    assert not bad, bad


def test_every_word_has_a_source_and_a_known_guard():
    """4: 出所の無い語は入らない（頭から思いついた語を足さない）。ガードは 3 種だけ。"""
    assert all(w.source for w in C.WORDS), [w.text for w in C.WORDS if not w.source]
    assert all(w.guard in C.GUARDS for w in C.WORDS), [w.text for w in C.WORDS if w.guard not in C.GUARDS]
    assert len({w.text for w in C.WORDS}) == len(C.WORDS), "同じ字面のレコードが 2 つある"


def test_labels_are_a_subset_of_the_product_labels():
    """★ 二重化した定数の番人（レビュー #8: 既に in/nin の分だけずれていた）── 部分集合で縛る。"""
    for k, v in C.LABELS.items():
        assert ailine._EXTRACT_CMP_LABELS.get(k) == v, (k, v, ailine._EXTRACT_CMP_LABELS.get(k))
    assert C.NUMERIC_CMPS is ailine.threshold.NUMERIC_CMPS, "数値の比較の集合が 2 つある"


# ── 5: 陰性対照（比較として読んではいけない句）────────────────────────────────

NEGATIVE_CONTROLS = [
    "会議の時間切って先に集計して",          # ★ レビュー #2: 『間』が数え語に当たり lt に化けた退行
    "納期は期日切って進めて",
    "この件切って次へ",
    "品名を区切って別の列に",
    "10日で締め切ってる案件",
    "報告は以上です。ご確認ください。",
    "先月は5000円でした。以下のとおり修正します",
    "売上は5000でした！以下のとおり修正して",          # ★ レビュー #9: 「。」以外の終止符
    "在庫は100だった\n以上ですがよろしく",              # ★ 改行も節の境界
    "報告は3件送りました！以上、よろしく",
    "予算を上回る努力をして",
    "締め切りを過ぎた案件",
    "退勤マイナス出勤の列を作って",
    "東京の行を抜き出して",
    "みかんの行とりんごの行だけを抽出して",
]


@pytest.mark.parametrize("task", NEGATIVE_CONTROLS)
def test_negative_controls_are_not_read_as_comparisons(task):
    r = C.read(task)
    assert not r.hit, (task, r)


# ── 5: 否定の活用の表（語幹 × 語尾）────────────────────────────────────────────

NEGATED_FORMS = [
    # 超え / を超えて
    ("残業時間が20時間を超えない人を抜き出して", "以下"),
    ("金額が5000を超えていない行を抜き出して", "以下"),     # ★ レビュー #1: 「い」1 文字で素通りしていた
    ("金額が5000を超えてない行", "以下"),
    ("金額が5000を超えません", "以下"),
    ("金額が5000を超えなかった行", "以下"),
    ("金額が5000を超えなければ", "以下"),
    # 下回 / 上回
    ("在庫数が10個を下回らない品番を抜き出して", "以上"),
    ("在庫数が10個を下回っていない品番", "以上"),
    ("在庫が10個を下回りません", "以上"),
    ("残業が20時間を上回らず済んだ人", "以下"),
    # 切っ / 過ぎ
    ("在庫数が10個切っていない品番", "以上"),
    ("退勤が22時を過ぎていない人", "以下"),
    # 以上 / 以下 / 未満
    ("金額が5000以上ではない行", "未満"),
    ("金額が5000以上じゃない行", "未満"),
    ("金額が5000以上になっていない行", "未満"),
    ("金額が5000以上ありません", "未満"),
    ("金額が5000以下でない行", "超"),
    ("金額が5000未満ではない行", "以上"),
    # マイナス
    ("在庫数がマイナスになっていない行を教えて", "以上"),
]


@pytest.mark.parametrize("task,inverted", NEGATED_FORMS)
def test_a_negated_comparison_is_asked_back_not_flipped(task, inverted):
    """① 否定は**反転させず聞き返す** ── 黙って反転させるのは直している事故と同じ形。
       ★ 過検出の害は小さく（聞き返すだけ）、過小検出だけが逆の実行を生む ── だから表で網羅する。"""
    r = C.read(task)
    assert r.cmp is None and r.hit and r.negated, (task, r)
    assert inverted in r.ambiguous and "否定" in r.ambiguous, r.ambiguous


@pytest.mark.parametrize("task,want", [
    ("残業時間が20時間超えてる人", "gt"),          # 「てる」は否定ではない
    ("在庫数が10個切ってるの教えて", "lt"),
    ("金額が5000以上ずつ足して", "gte"),           # 「ずつ」は否定ではない
    ("金額が5000を超えています", "gt"),
])
def test_conjugations_that_are_not_negation_stay_clean(task, want):
    r = C.read(task)
    assert (r.cmp, r.negated) == (want, False), (task, r)


# ── 引き方の番人（凍結した予測）─────────────────────────────────────────────────

def test_longest_match_wins_at_the_same_place(monkeypatch):
    """同じ位置から始まる 2 語は**長い方が勝つ**（「超え」が「超えない」を横取りしない）。

    ★ 正直に書く: 出荷している辞書には同じ位置から始まる 2 語が**無い**ので、この規則は今日の
      辞書では発火しない（変異試験で緑のままだった＝在っても鳴らない番人）。規則が要るのは
      **否定形を語として登録する日**（反転を定義する日）── だからその日の形の辞書を検体で作る。
    """
    extra = C.Word("超えない", "lte", C.NUM_BEFORE, source="検体用（この試験の中だけ）",
                   example="金額が5000を超えない行")
    monkeypatch.setattr(C, "WORDS", C.WORDS + (extra,))
    r = C.read("金額が5000を超えない行を抜き出して")
    assert (r.cmp, r.word, r.negated) == ("lte", "超えない", False), r
    # ★ 出荷の辞書では「を超える」（先に始まる）が「超え」（重なる）を畳む ── 語は 1 つだけ報告される
    r2 = C.read("金額が5000を超える行を抜き出して")
    assert (r2.cmp, r2.word) == ("gt", "を超える"), r2


@pytest.mark.parametrize("task,words", [
    ("金額が3000以上5000未満の行を抜き出して", ["以上", "未満"]),
    ("金額が1000以上で数量が10以下の行を抜き出して", ["以上", "以下"]),      # 2 列の条件も同じ断り
    ("数量が10より多く100より少ない行", ["より多く", "より少ない"]),
    ("金額が5000未満ないし3000以上", ["未満", "以上"]),             # ★ レビュー #10:「ないし」は否定でない
    ("金額が3000以上5000未満ではない行", ["以上", "未満"]),          # ★ 曖昧を否定より先に見る
])
def test_two_numeric_comparisons_are_refused_and_both_words_are_named(task, words):
    """② 数値の比較が 2 つ ── 黙って片方を捨てない。断り文は**可変部**で検査する（恒真にしない）。
       ★ 範囲か 2 列かは列名の解決が要るのでここでは決めない ── 断り文は両方を言う（レビュー #4）。"""
    r = C.read(task)
    assert r.cmp is None and r.hit and not r.negated, r
    assert _shown(r.ambiguous) == words, (r.ambiguous, _shown(r.ambiguous))
    assert "範囲" in r.ambiguous and "2 つの列" in r.ambiguous, r.ambiguous


@pytest.mark.parametrize("task,want,word", [
    ("金額が1000以上で部門が営業と同じ行の備考に「○」を付けて", "gte", "以上"),
    ("部門が営業と同じで金額が1000以上の行", "gte", "以上"),       # 同じ節なら並びに依らず数値の側
])
def test_a_numeric_comparison_beside_an_equality_is_two_conditions_not_a_range(task, want, word):
    """★ 曖昧と読むのは**数値の比較が 2 つ**だけ ── 数値＋等しい は 2 条件の依頼で、
       比較は数値の側（2 組目は実表の値で読む既存の道・test_two_conditions_are_not_half_done）。"""
    r = C.read(task)
    assert (r.cmp, r.word, r.ambiguous) == (want, word, ""), r


@pytest.mark.parametrize("sep", ["。", "！", "!", "？", "?", "．", ".", "\n"])
def test_numeric_priority_does_not_cross_a_sentence_boundary(sep):
    """★ レビュー #3: 「…を含む行を抜き出して。ちなみに売上は5000以上です」で後ろの雑談が
       主文の contains を奪っていた ── 数値優先は**同じ節の中**だけ。
       ★ レビュー #9: 節の境界は「。」だけではない（旧版からの癖）。"""
    r = C.read(f"備考に東京を含む行を抜き出して{sep}ちなみに売上は5000以上です")
    assert (r.cmp, r.word) == ("contains", "を含む"), (sep, r)


@pytest.mark.parametrize("task,want", [
    # ★ 比較が**1 つだけ**の文を使う ── 2 つ在る文は「曖昧」の規則が先に返してしまい、
    #   この規則自身が効いているかを測れない（変異試験で緑のままだった＝別の規則に守られた番人）。
    ("売上が10万以上ないし相当額の行を抜き出して", "gte"),
    ("金額が5000を超えないし方法がない", "gt"),
])
def test_a_negation_word_is_not_read_inside_naishi(task, want):
    """★ レビュー #10: 「ないし」は『または』── 否定ではない。
       ★ 正直に書く: 「超えないしくみ」のような語も同時に否定から外れるが、**実文が無く
         どちらの向きにも測っていない**（出たら測って決める）。"""
    r = C.read(task)
    assert (r.cmp, r.negated) == (want, False), (task, r)


def test_the_ambiguous_check_runs_before_the_negation_check():
    """★ レビュー #10（順序）: 比較が 2 つ在る文では、**2 つ在る事実**を先に言う。
       順を逆にすると「5000未満ないし3000以上」が片方の語の否定として返り、
       比較が 2 つ在ることを黙る ── どう書き直すかを決める材料が消える。
       ★ 番人は**中身**で縛る（字面の順でなく）── この 2 文は否定の語も比較 2 つも含む。"""
    for task in ("金額が5000未満ないし3000以上", "金額が3000以上5000未満ではない行"):
        r = C.read(task)
        assert not r.negated and len(_shown(r.ambiguous)) == 2, (task, r)


def test_two_non_numeric_words_keep_the_old_first_match_until_measured():
    """保留（実文が無い）: 「を含む」と「と同じ」が並ぶ回は旧来どおり先に出た語。
       ★ 発火条件: 実文が 1 件出た日に、範囲と同じ線で断るかを測ってから決める。"""
    r = C.read("備考に東京を含む行で状態が完了と同じもの")
    assert (r.cmp, r.word) == ("contains", "を含む"), r


def test_the_same_comparison_twice_is_not_ambiguous():
    r = C.read("金額が5000以上の行を、以上で抜き出して")
    assert r.cmp == "gte" and not r.ambiguous, r


def test_no_match_reads_nothing_and_says_so():
    r = C.read("在庫数が10を切っていない品番")     # ★ 辞書の漏れ ── 引けない回は決めない
    assert r == C.Reading(), r
    assert C.read("") == C.Reading() and C.read(None) == C.Reading()


@pytest.mark.parametrize("task,want", [
    ("単価が1000より高い行", "gt"),
    ("20件を超えた分だけ抜き出して", "gt"),
    ("金額が5000以上の行", "gte"),
])
def test_clean_readings_are_unchanged(task, want):
    assert ailine.extract_cmp_from_task(task) == want


# ── 3: 語彙の共食いの凍結（402 句）─────────────────────────────────────────────

def _now() -> list:
    frozen = json.loads(FREEZE.read_text(encoding="utf-8"))
    out = []
    for row in frozen["readings"]:
        r = C.read(row["phrase"])
        out.append({**row, "cmp": r.cmp, "word": r.word, "hit": r.hit,
                    "negated": r.negated, "ambiguous": bool(r.ambiguous)})
    return frozen, out


def test_the_frozen_readings_of_the_sweep_phrasings_do_not_move():
    """★ 辞書を変えた前後で、凍結した句の読みが**1 件も動かない**（語彙の共食いの番人）。動くなら
       意図した件だけのはずで、AILINE_REGEN_COMPARE_FREEZE=1 で再生成し、git diff を人が読んでから
       commit する（公開面の凍結と同じ作法）。雑音床 0（純関数）。
       ★ この凍結は**共食い専用** ── 比較を含む句が 5% しか無く、否定や断片ガードの番人にはならない
         （レビュー #5 で測った）。そちらは上の陰性対照と活用の表が受け持つ。"""
    frozen, now = _now()
    if REGEN:
        FREEZE.write_text(json.dumps({**frozen, "readings": now}, ensure_ascii=False, indent=1) + "\n",
                          encoding="utf-8")
        pytest.skip("凍結を再生成した ── git diff を読むこと")
    moved = [(a["phrase"], {k: a[k] for k in ("cmp", "word", "hit", "negated", "ambiguous")},
              {k: b[k] for k in ("cmp", "word", "hit", "negated", "ambiguous")})
             for a, b in zip(frozen["readings"], now) if a != b]
    assert not moved, f"{len(moved)} 件の読みが動いた（意図した件なら凍結を再生成）: {moved[:5]}"
    assert len(frozen["readings"]) >= 400, "凍結の句が減っている"


# ── 呼び出し側の規則（EXTRACT / SET_WHERE / フォルダ抽出が同じ器官を同じ規則で呼ぶ）──────

def _meta():
    return {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額", "印", "備考"]},
            "header_rows": {"Sheet": 1}}


def _extract(args, task):
    return ailine.verify_dsl_args("EXTRACT", args, _meta(), task=task)


def _set_where(args, task):
    return ailine.verify_dsl_args("SET_WHERE", args, _meta(), task=task)


_EX = {"col": "金額", "cmp": "gt", "value": 5000}
_SW = {"col": "印", "cond_col": "金額", "cmp": "gt", "cond_value": 5000, "value": "◎"}


@pytest.mark.parametrize("task,snippet", [
    ("金額が5000を超えない行", "否定"),
    ("金額が5000を超えていない行", "否定"),                 # ★ レビュー #1 の実害（LLM の正解を機械が壊す形）
    ("金額が3000以上5000未満の行", "2 つあります"),
])
def test_both_deciders_refuse_the_same_requests(task, snippet):
    """★ 1 本の試験で 2 経路を縛る（系譜「1 本の試験で全経路を縛る」）── 字面のカウントでなく振る舞い。
       フォルダ抽出は cmd_run_folder が重いので、そちらは下の配線の数で見る。"""
    ok1, _r, _i, err1 = _extract(dict(_EX), task + "を抜き出して")
    ok2, _r, _i, err2 = _set_where(dict(_SW), task + "の印に『◎』を付けて")
    assert not ok1 and snippet in err1, (ok1, err1)
    assert not ok2 and snippet in err2, (ok2, err2)


def test_a_correct_llm_reading_of_a_negated_request_is_not_overwritten():
    """★ レビュー #1 の実害そのもの: 「超えていない」で LLM が正しく lte を返した回を、
       機械が gt に上書きして**通していた**。いまは聞き返す（機械が勝つのは否定でない時だけ）。"""
    ok, r, _i, err = _extract({"col": "金額", "cmp": "lte", "value": 5000},
                              "金額が5000を超えていない行を抜き出して")
    assert not ok and "否定" in err, (ok, r.get("cmp"), err)


def test_a_fragment_in_the_preamble_does_not_break_a_contains_request():
    """★ レビュー #2 の実害: 「会議の時間切って、備考に至急って…」が lt に化けて断られていた。"""
    ok, r, _i, err = _extract({"col": "備考", "cmp": "contains", "value": "至急"},
                              "会議の時間切って、備考に至急って書いてある行を抜き出して")
    assert ok and r["cmp"] == "contains", (ok, err, r)


def test_extract_does_not_take_the_llm_inequality_when_the_dictionary_is_silent():
    """④ 引けなかった回に黙って LLM に負けない ── 三項（依頼／宣言／実体）の依頼側が欠けている。"""
    ok, r, _i, err = _extract({"col": "金額", "cmp": "lt", "value": 10}, "金額が10を切っていない行を抜き出して")
    assert not ok and "比較の語" in err and "未満" in err, (ok, err)


def test_extract_without_a_task_still_takes_the_dsl_as_is():
    """★ 依頼文が無い経路（DSL 直渡し・ゴールデン）は触らない ── 接地する相手が無い。"""
    ok, r, _i, err = _extract({"col": "金額", "cmp": "lt", "value": 10}, "")
    assert ok and r["cmp"] == "lt" and r["value"] == 10.0, (ok, err, r)


def test_extract_eq_with_a_number_is_not_touched_by_the_rule():
    """規則 ④ は**数値の比較**だけ ── eq は辞書に無くても通る（名指しの抽出の道）。"""
    ok, r, _i, err = _extract({"col": "金額", "cmp": "eq", "value": 5000}, "金額が5000の行を抜き出して")
    assert ok and r["cmp"] == "eq", (ok, err, r)


def test_set_where_does_not_take_the_llm_inequality_when_the_dictionary_is_silent():
    ok, r, _i, err = _set_where({**_SW, "cmp": "gte"}, "金額が5000の行の印に『◎』を付けて")
    assert not ok and "比較の語" in err, (ok, err)


def test_set_where_with_a_dictionary_word_is_unchanged():
    ok, r, _i, err = _set_where({**_SW, "cmp": "gte"}, "金額が5000以上の行の印に『◎』を付けて")
    assert ok and r["cmp"] == "gte" and r["cond_value"] == 5000.0, (ok, err, r)


def test_the_three_deciders_share_one_organ():
    """★ 片配線の番人: 比較を決める 3 経路（EXTRACT／SET_WHERE／フォルダ抽出）が全部
       compare_words.read を呼び、旧辞書（_EXTRACT_CMP_WORDS）が本体に残っていない。
       ★ 数は**等号**（レビュー #7: `>= 7` は呼び出しを 2 つ消しても緑だった）。
         呼ぶ場所は 7 ＝ 在否の門 3（条件つき書換／1 セル／名指しの読み直し）＋ 決定点 3
         （EXTRACT は 1 回の読みを名指しの門と決定点の 2 箇所で使う）＋ 包み 1。"""
    assert count_in_product("compare_words.read(") == 7, count_in_product("compare_words.read(")
    assert count_in_product("_EXTRACT_CMP_WORDS") == 0, "旧辞書が本体に残っている（辞書が 2 つになる）"
    assert count_in_product("compare_words.unconfirmed(") == 3, "規則 ④ が 3 経路に配線されていない"
    assert count_in_product("if _cmp_read.ambiguous:") == 3, "否定・曖昧の断りが 3 経路に配線されていない"


# ── ★ 実機で打って見つけた片配線（2026-09-15・関数の層では見えない）──────────────

def test_a_dry_plan_whose_steps_all_fail_does_not_claim_success(tmp_path, monkeypatch, capsys):
    """★★ 「売上が800以上1200未満の行を抜き出して」の `--dry` は全段が「× 未対応」なのに
       **exit 0 ＋ `"ok": true`** を返していた（履歴にも成功として残る）。
       ★ 同じ断りが単発の経路では exit 3 ── 同じ入力に 2 通りの返事（系譜「出力と終了コードも
         片配線する」）。★ 見つけ方は**実機で打って画面を読む**しかなかった ── 関数の層の
         検体 82 本は 1 本も鳴らなかった。
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from test_golden_transcripts import _book, _isolate, _run_main   # noqa: E402
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path, [["商品", "売上"], ["a", 900], ["b", 1300]])
    # ★ 段は**別々**にする ── 同じ args を 2 つ並べると道具が畳んで単発の経路へ落ち、
    #   計画の経路を測れない（初版はこれで空振りした）。
    plan = [{"op": "EXTRACT", "args": {"col": "売上", "cmp": "gte", "value": 800}},
            {"op": "EXTRACT", "args": {"col": "売上", "cmp": "lt", "value": 1200}}]
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"plan": plan})
    rc, out = _run_main(["run", str(book), "売上が800以上1200未満の行を抜き出して", "--dry"], capsys)
    assert "未対応" in out, out
    assert rc == 3, f"全段が未対応なのに exit {rc}（0 は『成功』の意味 ── 自動化が成功と読む）\n{out}"


def test_a_dry_plan_that_can_run_still_exits_zero(tmp_path, monkeypatch, capsys):
    """★ 陰性対照 ── 通るプレビューは今までどおり 0（直した所が別の所を壊していない）。"""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from test_golden_transcripts import _book, _isolate, _run_main   # noqa: E402
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path, [["商品", "売上"], ["a", 900], ["b", 1300]])
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "SORT", "args": {"col": "売上", "order": "desc"}},
                                  {"op": "BOLD", "args": {"target": "row:1"}}]})
    rc, out = _run_main(["run", str(book), "売上で降順に並べ替えて見出しを太字に", "--dry"], capsys)
    assert rc == 0, out
    assert "未対応" not in out, out
