# -*- coding: utf-8 -*-
"""依頼が名指しした**操作の種類**と、実行した操作が食い違っていないか。

★★ なぜ在るか（2026-09-07・外部の検品が最重の所見として拾った）:

    依頼   「ヤマノ食品の行を**削除して**」
    実行   操作:**抽出** → 新しいシートを作り、元の 6 行はそのまま
    出力   **✓ 機械検証済み**

  事後条件は「抽出として正しいか」を確かめるので通る ── **宣言と実体は一致していて、
  依頼だけが落ちている**。判定に要る三項のうち、また依頼が見られていなかった。

★ 既存の関所（`residue.unaccounted_request_words`）は**列名**しか見ないので、
  『削除』のような**動詞**は拾えない。ここはその隣を受け持つ。

★★ 何を見るか ── op 名の一致ではなく、**効果の種類**を見る:

      依頼文が「取り除く」系の op の語彙に当たっている
      かつ 実行した op が「取り除く」を書かない       → ✓ を出さない

  ★ op 名で見ると誤爆する。実測（4,538 件の実走行）:

      COMPUTE_COLUMN ← 「列を追加して」が ADD_COLUMN に当たる  x22  ← 上位下位の重なり
      DEDUP          ← 「行を消して」が DELETE_ROWS に当たる    x14  ← ★ 本物
      EXTRACT        ← 「削除して」                             x 2  ← ★ 本物

    効果の種類で見ると上の x22 は消え、**16 件 0.35% が残り、全部が同じ家系**だった
    （DEDUP も EXTRACT も新しいシートを作るだけで、元の行を取り除かない）。

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（既存の関所と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations


#: 「消す」と読める**裸の動詞**（op の照合語彙は「行を消して」の形しか持たないため）。
BARE_REMOVALS = ("消して", "削除して", "消す", "削除する", "取り除いて", "取り除く")

#: 数値書式で頼まれがちだが、この道具が**持っていない**書き方（語 → 人に見せる名前）。
#: ★ 持っているのは桁区切りだけ（FormatThousands）。
UNSUPPORTED_FORMATS = {
    "円マーク": "通貨記号（¥）", "￥": "通貨記号（¥）", "¥": "通貨記号（¥）",
    "通貨": "通貨記号（¥）", "パーセント": "百分率（%）", "％": "百分率（%）",
    "%": "百分率（%）", "小数": "小数点以下の桁数", "年月日": "日付の書き方",
}


def format_asked_but_not_supported(task: str, column_names=()) -> str | None:
    """依頼文が**この道具に無い書き方**を名指ししているなら、その名前を返す。

    ★★ なぜ在るか（2026-09-07・外部の査定が false ✓ として拾った）:

        依頼   「金額に**円マーク**を付けて」
        実行   数値書式 書式:thousands（桁区切り）→ 実ファイルは `#,##0`（¥ 無し）
        出力   **✓ 機械検証済み**

      ★ `style != "thousands"` を弾く番人は**在った**。だが LLM は「円マーク」を
        持っている書式へ**正規化して**返すので、宣言だけ見ていると素通りする。
        ── 三項（依頼・宣言・実体）のうち、また**依頼**が見られていなかった。

    ★ 列名に当たる語では鳴らさない（「日付の列に桁区切りを付けて」の『日付』は
      書式の指定ではなく対象）── 今日の他の判定と同じ**結び先を見る**作法。
    """
    text = task or ""
    cols = {str(c) for c in (column_names or ()) if c}
    for word, label in UNSUPPORTED_FORMATS.items():
        if word in text and word not in cols:
            return label
    return None


#: 「消す」意味の語 ── 値として書き込むと、literal で『空』と書いてしまう。
ERASERS = ("空", "空欄", "クリア", "未入力", "なし", "ブランク", "空白")


def why_not_a_value(value, column_names, sheet_names, op_words) -> str | None:
    """その語を**書き込む値として採ってはいけない**なら、その理由を返す（採れるなら None）。

    ★★ なぜ在るか（2026-09-07・Namakoo「値なのか操作なのか、属性なのかを判別出来るなら
      『』はいらない」）: 値を「」で囲ませているのは小型モデルの限界への回避策だった。
      ★ 実測（未見 20 本）: いまの製品のモデル(no thinking)は 17/20 まで取れる。
        外した中身は「操作を値と読む」「列名を値と読む」で、**そこは機械がタダで止められる**。

    ★ ここは**拒否だけ**をする（取り出しはモデルの仕事）── 役割を混ぜない。
    ★ 判定を広げない: 列名・シート名・操作の語・消す語、の 4 つだけ。
      「それ以外は値」と決めるのはこの関数ではなく、呼び出し側の段（空の列か等）。
    """
    v = str(value or "").strip()
    if not v:
        return "値が空です"
    if v in {str(c) for c in (column_names or ()) if c}:
        return f"『{v}』は列の名前です（書き込む値ではありません）"
    if v in {str(s) for s in (sheet_names or ()) if s}:
        return f"『{v}』はシートの名前です（書き込む値ではありません）"
    if any(w and w in v for w in (op_words or ())):
        return f"『{v}』は操作の名前を含みます（書き込む値ではありません）"
    if v in ERASERS:
        return f"『{v}』は消す操作です（その文字を書き込むことはできません）"
    return None


def removal_asked_but_not_done(task: str, op: str, vocab_by_op: dict,
                               removes: dict) -> list:
    """「取り除く」を頼まれたのに取り除かなかったなら、当たった語を返す（無ければ空）。

    task        … 依頼文
    op          … 実行した操作
    vocab_by_op … {op: その op の照合語彙}（`_op_match_pool` の結果）
    removes     … {op: その op が「取り除く」を書くか}（書き込み様式の登録簿から）

    ★ 実行した op **自身の語彙**が依頼文に当たっているなら、食い違いとは言わない
      （「重複行を削除して」で DEDUP が選ばれた回に、DEDUP の語彙も当たっていれば
      利用者はその操作を名指ししている）。
    """
    text = task or ""
    if not text or removes.get(op):
        return []                       # 取り除く op を実行したなら、食い違わない
    if any(p for p in (vocab_by_op.get(op) or ()) if p and p in text):
        return []                       # 実行した op 自身が名指しされている
    hits = []
    for other, phrases in (vocab_by_op or {}).items():
        if other == op or not removes.get(other):
            continue
        hits += [p for p in (phrases or ()) if p and p in text]
    if not hits:
        # ★★ 2026-09-07: **裸の動詞**まで見る。op の照合語彙は「行を消して」の形しか
        #   持たないので、「田中さんの分だけ残して**他は消して**」が素通りしていた
        #   （外部の査定が false ✓ として拾った 2 件目 ── 抽出に化けて ✓ が出ていた）。
        #   ★ 広げる前に測った: 4,556 件の実走行で 0.37% → 0.70%。増えた 15 件は
        #     すべて DEDUP（新シートを作るだけで元の重複は残る）＝既に本物と数えた家系。
        #     **新しい誤爆は 0 件**だったので広げる。
        hits = [w for w in BARE_REMOVALS if w in text]
    return list(dict.fromkeys(hits))
