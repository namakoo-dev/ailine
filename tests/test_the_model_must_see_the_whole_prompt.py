# -*- coding: utf-8 -*-
"""モデルに渡したつもりの材料が、本当に窓に入っているか。

★★ なぜ在るか（2026-09-09・第二の脳の HTTP 400 を追っていて ailine 側に見つけた）:

  ollama は文脈の窓を越えたプロンプトを **エラーにしない**。削るのでもない ──
  **窓の半分に落として、黙って生成する**。実測（qwen2.5-coder:7b・num_ctx 8192）:

      真の長さ 8,126 tok → 読まれた 8,126   ○ 無傷
      真の長さ 9,126 tok → 読まれた 4,098   ★ 崖
      語彙外段 26,574 tok → 読まれた 4,098  ★ 本番でこれが起きていた

  消えるのは**先頭から**（印を 3 か所に置いて確認: 先頭×／中ほど×／末尾○）。
  つまり語彙外段のモデルは CONTRACT（規約）もカタログの前 8 割も一度も見ずに
  Basic を書いていた。★ 例外は出ず、テストは全部緑のままだった。

★★ この試験は 2 層でできている:

  ①（ollama 不要・CI で毎回）実プロンプトの**文字数**を、実測した密度で
    トークンに換算し、その段の窓に収まるかを見る。★ 「窓を越えた」を主張する
    のではなく、**「最後に実測した時より伸びた＝測り直せ」**と言う番人。
  ②（要 ollama・@pytest.mark.local）密度の仮定そのものを実機で測り直す。

★ ①だけだと密度の仮定が古びても気づけない ── だから②を隣に置く。
"""
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))
import ailine
from ailine_core.prompt_window import (describe_the_loss, fewest_tokens_this_can_be,
                                       prompt_chars, tokens_the_model_never_read)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _product_source import count_in_product, window_around  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

#: 2026-09-09 に実機で測った密度（字/トークン）。★ 段ごとに違う ──
#: 翻訳層は日本語の散文が主で 1.60、語彙外段は Basic のコードが主で 2.56。
#: ★ 「同じ 1 つの比率」で済ませると、どちらかを必ず読み違える。
MEASURED_CHARS_PER_TOKEN = {"翻訳層": 6623 / 4153, "語彙外段": 68043 / 26574}

#: 生成に使う分の余白（num_predict）。プロンプトと足して窓に収まる必要がある。
NUM_PREDICT = {"翻訳層": 300, "語彙外段": 1600}

#: 換算の粗さに対する安全率。★ 密度は段の中身が変われば動くので、ぴったりでは持たない。
SAFETY = 1.10


def _translation_messages():
    meta = {"headers": {"見積": ["品名", "単価", "数量", "金額"]}}
    return ailine.build_translation_messages("単価の大きい順に並べ替えて", meta)


def _freeform_messages():
    catalog, _files = ailine.load_helpers(REPO_ROOT / "src" / "ailine" / "helpers")
    system = ailine.CONTRACT + ailine.load_refs(REPO_ROOT / "src" / "ailine" / "refs") + catalog
    return [{"role": "system", "content": system},
            {"role": "user", "content": "タスク:\n単価の大きい順に並べ替えて\n\n"
                                        "`Sub Run(oDoc As Object)` を1つだけ書け。コードのみ。"}]


@pytest.mark.parametrize("stage, messages_of, window", [
    ("翻訳層", _translation_messages, "NUM_CTX"),
    ("語彙外段", _freeform_messages, "NUM_CTX_FREEFORM"),
])
def test_each_prompt_still_fits_the_window_it_is_sent_with(stage, messages_of, window):
    """語彙を足してプロンプトが伸びた時、窓を越える前に CI で気づく。"""
    num_ctx = getattr(ailine, window)
    room = num_ctx - NUM_PREDICT[stage]
    chars = prompt_chars(messages_of())
    estimated = chars / MEASURED_CHARS_PER_TOKEN[stage]
    assert estimated * SAFETY <= room, (
        f"{stage} のプロンプトが窓に収まらなくなった: {chars:,} 字 ≒ {estimated:,.0f} tok"
        f"（安全率込み {estimated * SAFETY:,.0f}） > 使える {room:,} tok（窓 {num_ctx:,}）\n"
        f"★ ollama は例外を出さず**窓の半分に落として黙って生成する**。\n"
        f"★ {window} を広げるか、渡す材料を減らすこと。広げる時は VRAM も測ること"
        f"（32768 で既に 8% が CPU に溢れている・RTX 4060 8GB 実測 2026-09-09）。")


def test_the_guard_can_tell_read_from_unread():
    """番人が『読まれた』と『読まれていない』を区別できることを、両方向で確かめる。

    ★ 片方向だけの試験は恒真になりうる ── この repo で 2 回踏んだ。"""
    msgs = [{"role": "system", "content": "あ" * 40000}]      # 下限 10,000 tok
    assert fewest_tokens_this_can_be(msgs) == 10000

    # 全部読まれた側: 取りこぼしの証拠は出ない
    assert tokens_the_model_never_read(msgs, 26000) == 0
    assert describe_the_loss(msgs, 26000, 32768) == ""

    # 窓の半分に落ちた側: 証拠が出る
    assert tokens_the_model_never_read(msgs, 4098) == 10000 - 4098
    said = describe_the_loss(msgs, 4098, 8192)
    assert "読んでいない" in said and "5,902" in said


def test_a_missing_count_is_not_called_an_accident():
    """prompt_eval_count が無い／数でない時に事故だと言わない（分からないと事故は別）。"""
    msgs = [{"role": "system", "content": "あ" * 40000}]
    for absent in (None, "4098", -1, {}):
        assert tokens_the_model_never_read(msgs, absent) == 0
        assert describe_the_loss(msgs, absent, 8192) == ""


def test_both_ollama_paths_go_through_the_single_guarded_door():
    """通信の診断と窓の検算が **1 か所** に畳まれていることを静的に確かめる。

    ★ 〈片配線〉への番人 ── 以前は通常生成にだけ診断が在り、翻訳層は try すら
      無かった。「両方に足す」で直すと次に足す人がまた片方だけにする。"""
    doors = count_in_product('urllib.request.Request(ollama_url("/api/chat")')
    assert doors == 1, f"/api/chat を叩く口が {doors} 箇所ある（1 箇所に畳むこと）"
    for fn in ("def ollama_generate(", "def ollama_generate_json("):
        # ★ 本体の場所を決め打ちしない ── 分割で実装が動いても番人が空振りしない
        body = window_around(fn, after=1500).split(chr(10) + "def ", 1)[0]
        assert "ask_ollama(" in body, f"{fn} が ask_ollama を通っていない"


@pytest.mark.local
@pytest.mark.parametrize("stage, messages_of, window", [
    ("翻訳層", _translation_messages, "NUM_CTX"),
    ("語彙外段", _freeform_messages, "NUM_CTX_FREEFORM"),
])
def test_the_machine_actually_reads_all_of_it(stage, messages_of, window):
    """実機で測り直す ── 上の密度の仮定が古びていないか。★ 要 ollama。"""
    import json
    import urllib.request
    msgs = messages_of()
    num_ctx = getattr(ailine, window)
    body = {"model": "qwen2.5-coder:7b", "messages": msgs, "stream": False,
            "options": {"temperature": 0, "num_predict": 1, "num_ctx": num_ctx}}
    req = urllib.request.Request(ailine.ollama_url("/api/chat"),
                                 data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        got = json.load(r)
    read = got.get("prompt_eval_count")
    assert not describe_the_loss(msgs, read, num_ctx), (
        f"{stage}: 実機で取りこぼしが出た（読んだ {read} / 窓 {num_ctx}）")
    density = prompt_chars(msgs) / read
    frozen = MEASURED_CHARS_PER_TOKEN[stage]
    assert abs(density - frozen) / frozen < 0.15, (
        f"{stage} の密度が凍結値からずれた: 実測 {density:.2f} 字/tok vs 凍結 {frozen:.2f}\n"
        f"★ MEASURED_CHARS_PER_TOKEN を測り直して更新すること（①の換算がこれに依存する）")
