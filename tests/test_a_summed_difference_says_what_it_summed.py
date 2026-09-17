# -*- coding: utf-8 -*-
"""足し算の差額は「何を足したか」を画面で言う（2026-09-17・盲検 3 体目）。

★★ 起きたこと（製造業の購買担当・初見・盲検）:
  標準単価と発注の単価を突き合わせた画面が、こう出した ──

      ⚠ エポキシ接着剤: +5200（A 10000 / B 4800）。明細シートで内訳を確認してください。

  エポキシは 9 月に 2 回発注していて、単価は 4800（標準どおり）と 5200（値上げ）。
  道具は**単価を足して** 10000 にし、標準 4800 と引いて +5200 と出した。
  買い手が知りたいのは「1 本あたり +400 円」で、+5200 は
  **「仕入先に文句を言いに行ったら恥をかく数字」**（買い手の言葉）。exit は 0。
  件数（A 2 件）は出力シートには在ったが、**画面には無かった**。
  ★ 同じ品目を月に 2 回買うのは普通なので、この誤読は毎月出る。

★★ これは **2026-08-24 第三波 S5 と同じ読めなさ**の別の形だった ──
  あの時は『A 186300 / B 0』の 0 が「金額 0」か「1 行も無い」か読めなかった。
  今度は『A 10000』が「1 件の 10000」か「2 件の合計」か読めない。
  ★ 直したのは**同じ関数**（side_pair）── 実装は 1 つ、呼び手は 2 つのまま。
    だからこの試験は**両方の読めなさを 1 本で縛る**（片方だけ直さない）。

★ 「明細を見てください」では足りなかった。買い手は ⚠ を「明細で内訳を見ろ」と読み、
  **数字そのものは正しい**と信じた。だから画面で**数の作り**を言う。
★ 雑音は足さない ── 1 件しか無い側は今までどおり黙る（全行に読む理由の無い語を増やさない）。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
from ailine_core.match import (  # noqa: E402
    KeyGroup, build_findings, side_pair, sums_more_than_one_row)


def _g(key, a_count, a_sum, b_count, b_sum, state):
    return KeyGroup(key_display=key, a_count=a_count, a_sum=a_sum,
                    b_count=b_count, b_sum=b_sum, diff=a_sum - b_sum,
                    state=state, a_rows=[], b_rows=[])


#: 事故そのもの（買い手の冊の数字をそのまま）。
EPOXY = _g("エポキシ接着剤", 2, 10000.0, 1, 4800.0, "+5200")
#: 1 件どうし（雑音を足していないことの陰性対照）。
SERVO = _g("サーボモータ 400W", 1, 62000.0, 1, 61000.0, "+1000")


def test_a_sum_of_several_rows_says_how_many():
    """★ 事故の形: 2 行を足した側は、件数と『計』を名乗ること。"""
    s = side_pair(EPOXY)
    assert "2件" in s, f"何件を足したのか画面から読めない: {s}"
    assert "計10000" in s or "計 10000" in s, f"合計であることを名乗っていない: {s}"


def test_a_single_row_stays_quiet():
    """★★ 陰性対照: 1 件しか無い側に語を増やさない。

    ★ ここが無いと「全部に『1件 計』と書く」直しでも上の試験が通る ──
      それは読む理由の無い語を全行に増やすだけで、事故は減らない。
    """
    s = side_pair(SERVO)
    assert "件" not in s, f"1 件しかない側に余計な語が増えた: {s}"
    assert "計" not in s, f"1 件しかない側に余計な語が増えた: {s}"


def test_zero_is_still_told_apart_from_absent():
    """★★ 2026-08-24 第三波 S5 の線を落としていないこと（同じ関数を直したので一緒に縛る）。"""
    absent = _g("甲社", 2, 186300.0, 0, 0.0, "A のみ")
    assert "なし（0 行）" in side_pair(absent), f"0 と『無い』が区別できない: {side_pair(absent)}"
    real_zero = _g("乙社", 1, 100.0, 1, 0.0, "+100")
    assert "なし" not in side_pair(real_zero), f"本物の金額 0 を『無い』と偽った: {side_pair(real_zero)}"


def test_the_summing_sides_are_named():
    """★ 「どちら側が足し算なのか」を 1 箇所で決める（呼び手が書き写さない）。"""
    assert sums_more_than_one_row(EPOXY) == ["A"]
    assert sums_more_than_one_row(SERVO) == []
    both = _g("両方", 3, 300.0, 2, 100.0, "+200")
    assert sums_more_than_one_row(both) == ["A", "B"]


def test_the_screen_says_this_is_not_a_per_item_difference():
    """★★ 事故の核心 ── 1 件あたりの差ではない、と**画面で**言うこと。"""
    findings = build_findings([EPOXY], {}, [], "発注台帳.xlsx", "標準単価表.xlsx")
    assert findings, "★ 所見が 1 件も出ていない（下の検査が素通りする）"
    text = " ".join(str(f.next_step) for f in findings)
    assert "1 件あたりの差ではありません" in text, (
        f"足し算の差額を、1 件あたりの差と読まれるまま出している: {text}")


def test_a_one_to_one_difference_is_not_hedged():
    """★ 陰性対照: 1 件どうしの差額には、その断りを付けない（付けたら全行が濁る）。"""
    findings = build_findings([SERVO], {}, [], "発注台帳.xlsx", "標準単価表.xlsx")
    assert findings, "★ 所見が 1 件も出ていない"
    text = " ".join(str(f.next_step) for f in findings)
    assert "1 件あたりの差ではありません" not in text, (
        f"1 件どうしの正しい差額にまで断りが付いた: {text}")
