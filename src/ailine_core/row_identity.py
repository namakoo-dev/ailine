"""行の中で成り立っている列どうしの等式を見つけ、操作で崩れたら知らせる（純ロジック）。

★★ 2026-08-31（Namakoo）:「金額が入れ替われば付随して関連するセルの内容も変えなければ
  いけない。しかもそれが複数の内容に影響する場合はそれらも踏まえて変更しないといけない」

  実測: 「丸和物流の単価とみどり建設の単価を入れ替えて」は**頼まれた 2 セルだけ**を
  正しく入れ替える。だが 金額（＝件数×単価）は**直値**なので取り残され、
      件数 12 × 単価 7200 = 86,400  なのに 金額 57,600 のまま
  という**表として矛盾した状態**になる。それでも「2 セルだけ動いた」は真実なので ✓ が出る。

★ 式で書かれていれば LibreOffice が再計算するので起きない。**直値で持っている派生列**
  だけで起きる ── 見た目は普通の数字なので、人は気づけない。

★ 直さない・**言う**。どう直すか（金額を計算し直すのか、単価を戻すのか）は人が決める。
  ここは「操作の前に成り立っていた等式が、後で崩れた」という**事実だけ**を返す。

★ 語も見出しも読まない ── **数だけ**を見る（sum_identity と同じ性質）。
"""

from __future__ import annotations

_OPS = (("×", lambda a, b: a * b), ("＋", lambda a, b: a + b),
         ("−", lambda a, b: a - b))
_MIN_ROWS = 3          # ★ 偶然の一致を拾わないための下限
_TOL = 1e-6


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def identities(rows: list) -> list:
    """全データ行で成り立つ (結果列, 演算, 左列, 右列) を返す（0 起点の列番号）。

    rows: データ行のリスト（見出しは含めない）。値は数値か、それ以外（None 扱い）。
    ★ 全行で成り立つものだけ ── 1 行でも外れたら等式ではない。
    ★ 左右が同じ列、結果が入力と同じ列、は除く（自明・恒真になる）。
    """
    # ★★ 2026-08-31（最初の実装が空振りした・測定器を先に疑って分かった）:
    #   合計行は 件数・単価 が空なので、その列に None が在るとして**全部捨てていた**。
    #   ★ 数だけを見るこの層は「合計行」を知らない ── 呼ぶ側が外す（下の drop_rows）。
    #     ここでは「**数が揃っていない行は無視する**」に緩める（行ごとに判定する）。
    nums = [[_num(v) for v in r] for r in rows]
    if len(nums) < _MIN_ROWS:
        return []
    width = min(len(r) for r in nums) if nums else 0
    out = []
    for t in range(width):
        for a in range(width):
            for bcol in range(width):
                if len({t, a, bcol}) != 3:
                    continue
                # ★ 3 つとも数が入っている行だけで判定する（合計行のような欠けた行は無視）。
                usable = [r for r in nums
                           if r[t] is not None and r[a] is not None and r[bcol] is not None]
                if len(usable) < _MIN_ROWS:
                    continue
                for label, fn in _OPS:
                    if all(abs(fn(r[a], r[bcol]) - r[t]) <= _TOL for r in usable):
                        # ★ 掛け算・足し算は交換法則で 2 回当たる ── 同じ等式を
                        #   2 行に出さない（人には同じことに見える）。
                        if label in ("×", "＋") and (t, label, bcol, a) in out:
                            break
                        out.append((t, label, a, bcol))
                        break
    return out


def _lost(before_rows: list, after_rows: list) -> list:
    """前に成り立っていた等式のうち、後で成り立たなくなったもの（共通の芯）。

    ★ broken / broken_after_insert が**同じ実装を通る**ようにここへ畳んだ
      （2 通りに書くと、片方だけ直す事故が起きる ── この repo で何度も踏んだ形）。
    """
    keep = identities(before_rows)
    if not keep:
        return []
    still = set(identities(after_rows))
    return [x for x in keep if x not in still]


def broken(before_rows: list, after_rows: list) -> list:
    """操作の前に成り立っていて、後で崩れた等式（無ければ空）。

    ★ 行数が変わる操作では**ここでは**比べない ── 増える側は broken_after_insert が
      受け持つ。素朴に比べると「行を足したら必ず ⚠」になる（鳴りすぎ）。
    """
    if len(before_rows) != len(after_rows):
        return []
    return _lost(before_rows, after_rows)


def broken_after_insert(before_rows: list, after_rows: list) -> list:
    """行が増える操作で、**増えた行を含めて**等式がまだ成り立つか。

    ★★ 2026-09-02: 行を足すと、既存の行では式で出ている列（税込金額など）が
      新しい行だけ空になる ── 宣言した値は正しいので `✓` は正しいが、
      **表としては辻褄が合っていない**。ここはそれを見る。
    ★ 鳴りすぎない理由: identities() は **3 つとも数が入っている行だけ**で判定するので、
      値を入れなかった新しい行は最初から無視される。鳴るのは
      「新しい行に数が揃っていて、しかも等式を満たさない」時だけ ── 見たい事故そのもの。
    """
    if len(after_rows) <= len(before_rows):
        return []
    return _lost(before_rows, after_rows)


def describe(broken_list: list, headers: list) -> str | None:
    """崩れた等式を 1 行にする（無ければ None）。"""
    if not broken_list:
        return None

    def _nm(i):
        return str(headers[i]) if 0 <= i < len(headers) and headers[i] else f"{i + 1}列目"

    parts = [f"『{_nm(t)}』＝『{_nm(a)}』{op}『{_nm(b)}』" for t, op, a, b in broken_list[:3]]
    more = f" ほか {len(broken_list) - 3} 件" if len(broken_list) > 3 else ""
    return ("この操作の前は " + "・".join(parts) + more
             + " が全行で成り立っていましたが、**成り立たなくなりました**"
               " ── 計算で出している列が、直した値に付いていっていません"
               "（どう直すかは人が決めることなので、直していません）")
