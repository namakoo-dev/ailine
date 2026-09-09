# -*- coding: utf-8 -*-
"""ヘルパは**呼び方**だけ見せる。中身は見せない。

★★ なぜ在るか（2026-09-09）: 語彙外段のプロンプトは 26,574 トークンあり、その
  **94%（24,800 tok）が .bas の実装本文**だった。そしてカタログの冒頭にはこう書いてある:

      ★ ヘルパの中身は絶対に書き写すな（SummaryTable 等が長くても）
      ## 定義済みヘルパ（★ 呼ぶだけ・再定義しない）

  ── **読ませないために 24,800 トークン渡していた**。呼ぶのに要るのは署名と引数の
  意味だけで、実装は要らない。

★ 窓の話と直結する: ollama は窓を越えると例外を出さず**半分に落として黙って生成する**
  ので、この 26,574 tok は実際には 4,098 tok しか読まれていなかった（CONTRACT も
  カタログの前 8 割もモデルは一度も見ていない）。
  ★ Namakoo 決裁（2026-09-09）:「ノート PC 想定でハードウェア性能への依存を避けたい」
    ── 窓を広げて解決しない。**プロンプトを床（8,192）に合わせる**。

★ 落とすものは 2 つだけ:
    ① 実装本文（Sub 〜 End Sub の中身）
    ② 保守する人間向けの経緯（★ で始まる行・`W3:` などの設計事情・その続き行）
  ★ 残すのは 1 行目の説明と、引数 1 個 1 行の意味 ── モデルが呼ぶのに要る情報。
  ★ ②を落とすのは容量のためだけではない: 内部の設計事情をモデルに流さない。
"""
from __future__ import annotations

import re

#: `'   headerRow  : 見出し行（…）` の形（引数 1 個 1 行の説明）。
_ARG_DOC = re.compile(r"^'\s{2,}(\w+)\s*:")
#: 保守する人間向け ── モデルには要らない行。
_FOR_MAINTAINERS = re.compile(r"★|^'\s*W\d|spike|ours")
_SUB = re.compile(r"^Sub\s+[A-Za-z_]\w*\s*\(")


def _doc_for_model(doc: list) -> list:
    """説明コメントから、モデルが呼ぶのに要る行だけを残す。"""
    kept = []
    for i, line in enumerate(doc):
        if _FOR_MAINTAINERS.search(line):
            continue
        if i == 0 and not _ARG_DOC.match(line):
            kept.append(line)                     # 1 行目＝何をするヘルパか
        elif _ARG_DOC.match(line):
            # ★ 長い括弧書きは設計事情が混ざるので落とす（引数の意味だけ残す）
            kept.append(line.split("（")[0].rstrip() if len(line) > 90 else line)
    return kept


def interfaces_only(bas_text: str) -> str:
    """`.bas` の原文から、**呼び方だけ**のカタログを作る。

    ★ 行継続（末尾 `_`）で 2 行に分かれた署名も 1 つとして拾う ── 落とすと
      「そのヘルパは存在しない」とモデルに伝わる（実在するのに使えなくなる）。
    """
    lines = bas_text.replace(chr(13), "").split(chr(10))
    out, i = [], 0
    while i < len(lines):
        if not _SUB.match(lines[i]):
            i += 1
            continue
        j, doc = i - 1, []
        while j >= 0 and lines[j].lstrip().startswith("'"):
            doc.insert(0, lines[j])
            j -= 1
        sig, k = [lines[i]], i
        while sig[-1].rstrip().endswith("_"):
            k += 1
            sig.append(lines[k])
        out += _doc_for_model(doc) + sig + ["End Sub", ""]
        i = k + 1
    return chr(10).join(out)


def names_in(catalog_or_bas: str) -> list:
    """カタログ／原文に現れる Sub 名（★ 絞った後も欠けていないかを番人が数える）。"""
    return re.findall(r"^Sub\s+([A-Za-z_]\w*)", catalog_or_bas.replace(chr(13), ""), re.M)
