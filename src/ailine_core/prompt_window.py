# -*- coding: utf-8 -*-
"""モデルが**読まなかった分**を、送った側で気づけるようにする。

★★ なぜ在るか（2026-09-09・第二の脳の HTTP 400 を追っていて ailine 側に見つけた）:

  ollama は文脈の窓（num_ctx）を越えたプロンプトを **エラーにしない**。
  少しずつ削るのでもない ── **窓の半分に一気に落として、黙って生成する**。
  実測（qwen2.5-coder:7b・num_ctx 8192）:

      真の長さ 8,126 tok → 読まれた 8,126   ○ 無傷
      真の長さ 9,126 tok → 読まれた 4,098   ★ 5,028 tok 消えた
      真の長さ 26,574 tok → 読まれた 4,098  ★ 22,476 tok 消えた

  消えるのは**先頭から**。印を 3 か所に置いて確かめた（先頭×／中ほど×／末尾○）。
  つまり語彙外段では、モデルは CONTRACT（規約そのもの）もヘルパカタログの前 8 割も
  一度も見ないまま Basic を書いていた。★ 例外は出ず、テストは緑のまま。

★★ どう検算するか ── **別の道で出した数と突き合わせる**:

  自分で数えたトークン数を自分で確かめても恒真になる（この repo で何度も踏んだ）。
  そこで、こちらは **文字数から出す下限**（「どんなに詰まっていてもこれ以下ではない」）
  だけを持ち、実測値は **ollama が返す prompt_eval_count** を使う。
  下限すら下回っていたら、読まれなかったことの証拠になる。

★ 限界（正直に）:
  これは「**窓の半分に落ちる**」形の取りこぼしを捕まえる道具で、
  「1 トークンでも欠けたら分かる」ものではない。下限は安全側に粗く取ってある。
"""
from __future__ import annotations

#: 1 トークンあたりの文字数の**下限**。実測（2026-09-09）:
#:     日本語の散文        1.35 字/tok
#:     翻訳層のプロンプト  1.60 字/tok
#:     Basic のカタログ    2.56 字/tok
#: いちばん薄い 1.35 に対して 4.0 は 3 倍近い余裕がある ── ★ 誤検知を出さない側に倒す。
#: （粗いのは承知の上。捕まえたいのは「半分に落ちる」という桁の事故なので、これで足りる）
CHARS_PER_TOKEN_FLOOR = 4.0


def prompt_chars(messages: list) -> int:
    """送るメッセージ全体の文字数。role や区切りの分は数えない（下限を下げる側で安全）。"""
    return sum(len(m.get("content") or "") for m in messages)


def fewest_tokens_this_can_be(messages: list) -> int:
    """このプロンプトが取りうる**最小の**トークン数。実測より必ず小さくなる。"""
    return int(prompt_chars(messages) / CHARS_PER_TOKEN_FLOOR)


def tokens_the_model_never_read(messages: list, prompt_eval_count) -> int:
    """モデルが読まなかったと**言い切れる**トークン数。0 なら取りこぼしの証拠は無い。

    ★ prompt_eval_count が無い／数でない場合は 0 を返す（分からないことを事故と呼ばない）。
    ★ 「プレフィックス・キャッシュが効くと小さく報告されるのでは」と疑って測ったが、
      **外れた** ── 同じ依頼を 3 回続けても 4,126 / 4,126 / 4,126 と一定だった
      （2026-09-09・qwen2.5-coder:7b）。誤検知の源が無いので、呼び出し側は
      この値を**断りの根拠**にしてよい。★ 疑いを晴らしてから強くした順序を残す。
    """
    if not isinstance(prompt_eval_count, int) or prompt_eval_count < 0:
        return 0
    missing = fewest_tokens_this_can_be(messages) - prompt_eval_count
    return missing if missing > 0 else 0


def describe_the_loss(messages: list, prompt_eval_count, num_ctx: int) -> str:
    """人に見せる 1 行。取りこぼしが無ければ空文字列。"""
    missing = tokens_the_model_never_read(messages, prompt_eval_count)
    if not missing:
        return ""
    return (f"★ プロンプトが文脈の窓に入っていない: 少なくとも {missing:,} トークンを"
            f"モデルは読んでいない（読んだ {prompt_eval_count:,} / 窓 {num_ctx:,}・"
            f"送った {prompt_chars(messages):,} 字）。"
            "\n  ★ 消えるのは**先頭**（規約やカタログの前半）。num_ctx を広げるか、"
            "渡す材料を減らすこと。")
