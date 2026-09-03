"""小さな共有述語 ── 値の素性・数値の表示・見出しからの列番号。

★ なぜ在るか（2026-09-03 に数えた）: 同じ数行の実装が、本体と ailine_core に
**写し取られて**いた ── `_is_number` は 7 箇所、`fmt_num` は 3 箇所、
`_column_index` は 2 箇所。どれも docstring で互いを指しており
（「total_row._is_number と同じ線」）、**重複を自覚したまま注記で済ませていた**。
この repo の系譜「二重化した経路は片配線が既定で起きる」の、静かな側の実例。

★ 処方は系譜どおり ── **両方直す**のではなく **1 関数に畳んで呼び出し側に持たせない**。
呼ぶ側は `from ailine_core.primitives import is_number as _is_number` の形で引く
（モジュール内の私的な別名として使えるので、呼び出し行を書き換えずに済む）。

★ ここに置くものの線: **引数だけで結果が決まる**小さな述語だけ。
モジュール定数を参照するもの・ブックを読むものは置かない
（畳める根拠が「参照する定数の値まで一致する」ことだったため）。

★ ここに**置かなかった**もの（意図的に別・畳んではいけない）:
  ・`_values_agree` … csv_quarantine は型込み等値で TOLERANCE 不使用、
    verify は許容誤差つき。**同じ名前で違う判断**なので、畳むと片方の規則が漏れる
  ・`_is_empty`（cellmap）と `_is_blank_cell`（total_row） … 前者は `strip()` で
    タブ・改行も空とみなし、後者は半角/全角スペースのみ。**挙動が違う**ので
    畳むには先に実害の有無を測る必要がある（2026-09-03 時点で未測定）
"""
from __future__ import annotations


def is_number(v) -> bool:
    """数値セルとして扱えるか。

    ★ bool は int のサブクラスだが数値としては扱わない（True/False の混入対策）。
    """
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def fmt_num(v) -> str:
    """数値表示: 整数値は小数点なしで（600.0 でなく 600）。数値でなければそのまま str。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else str(f)


def column_index(headers: list, name: str):
    """見出しの一覧から 1 起点の列番号を返す。見つからなければ None。

    ★ 同名の見出しが複数あるときの扱いは**呼び出し側が先に断る**
      （multifile.duplicate_header_names）。ここで最初の 1 本を返す挙動そのものは
      変えない ── 変えると位置の意味が経路ごとに食い違い、やる側と見る側で
      また別の嘘が生まれる。★ この注記は multifile._column_index の docstring から
      持ってきた（2026-09-03 に畳んだ時、説明だけ失わないように）。
    """
    try:
        return headers.index(name) + 1
    except ValueError:
        return None
