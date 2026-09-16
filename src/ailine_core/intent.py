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

import re

#: 「消す」と読める**裸の動詞**（op の照合語彙は「行を消して」の形しか持たないため）。
BARE_REMOVALS = ("消して", "削除して", "消す", "削除する", "取り除いて", "取り除く")

# ── 見るだけの依頼（2026-09-14・言い回し 120 件の盲検・誤配の家系②）──────────────────
#
# ★★ 「品番の重複がないか**見といて**」→ 重複除去、「〜がないか**調べて**」→ 重複除去。
#   確認の依頼が**行を減らす**操作に写されていた。既定は新しい冊に書くので原本は壊れないが、
#   出てくるのは「重複を除いた表」で、**どれが重複だったかは分からない**（頼んだことの逆側）。
# ★ 一方、確認語が**抽出**（読むだけ・新しいシート）に着いた依頼は盲検で 9/9 正だった。
#   だから規則は狭い: 確認語が在り、**変更の動詞が 1 つも無い**依頼を、減らす op に写さない。
# ★ 語の列挙で正しい ── 判定しているのは「人がどの動作を言ったか」で、動作は言葉でしか
#   分からない（`_REMOVAL_WORDS` と同じ理屈）。漏れた語は「今までどおり」になるだけで、
#   黙って別のことをする方向には壊れない。
CHECK_ONLY_WORDS = ("ないか", "あるか", "調べて", "調べといて", "確認して", "確認しといて",
                    "チェックして", "見といて", "見ておいて", "教えて", "探して")
#: 在れば「見るだけ」ではない（減らす・書く・変える意思が言葉に出ている）。
CHANGE_VERBS = BARE_REMOVALS + ("削除", "除いて", "除く", "取り除", "抜いて", "無くして",
                                "直して", "書いて", "入れて", "変えて", "にして", "埋めて",
                                "付けて", "塗って", "並べ")
#: 行や列を**減らす** op ── 見るだけの依頼をここへ写したら止める。
ROW_REDUCING_OPS = frozenset({"DEDUP", "DELETE_ROWS", "DELETE_COLUMN"})
#: 減らしてよいときの通る書き方（op ごと）。
REDUCING_EXAMPLES = {"DEDUP": "品番が同じ行を重複として除いて",
                     "DELETE_ROWS": "金額が0の行を削除して",
                     "DELETE_COLUMN": "備考の列を削除して"}


def check_only_request(task: str | None) -> str | None:
    """依頼が「見るだけ」なら、その確認語を返す（変更の動詞が 1 つでも在れば None）。"""
    text = task or ""
    if any(v in text for v in CHANGE_VERBS):
        return None
    for w in CHECK_ONLY_WORDS:
        if w in text:
            return w
    return None


def refuse_reducing_a_check(task: str | None, op: str, label: str) -> str | None:
    """見るだけの依頼が減らす op に写されていたら、断り文を返す（そうでなければ None）。

    ★ 「見つけるだけの操作は無い」と正直に言い、減らしてよいときの通る書き方を添える。
    """
    if op not in ROW_REDUCING_OPS:
        return None
    word = check_only_request(task)
    if word is None:
        return None
    return (f"『{word}』は見るだけの依頼と読めますが、『{label}』は行や列を減らす操作です"
            "（見つけるだけの操作はありません）。減らしてよければ"
            f"「{REDUCING_EXAMPLES.get(op, '…を削除して')}」のように頼んでください")

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


#: この道具が **やらないこと**（語 → 人に見せる名前）。属性（文字色・斜体…）と、
#: 表そのものでない設定（ウィンドウ枠・印刷範囲…）の両方を持つ。
#:
#: ★★ なぜ在るか（2026-09-09・語彙外の台が初回で拾った**決定論的な false ✓**）:
#:
#:     依頼   「品名の**文字色**を赤にして」
#:     実行   操作:背景色 対象:col:品名 色:red
#:     出力   **✓ 機械検証済み**（3 回とも同じ・運ではない）
#:
#:   既存の 2 つの関所はどちらも**原理的に**この形が見えなかった:
#:     residue.unaccounted_request_words … 実在する**列名**しか見ない（『文字色』は列名でない）
#:     op_effect_mismatch                … 文字色も背景色も同じ**効果の種類**（書式のみ）
#:   ★ 粒度が 1 段足りなかった ── 見るべきは列でも効果でもなく **属性**。
#:
#: ★ これは「持っていない」を宣言する名簿なので、**広すぎても実害は『断りすぎ』で済む**
#:   （許可の名簿とは非対称 ── 狭すぎる許可は、正しい依頼を黙って別物に変える）。
#: ★ 持っている能力を名簿で殺さないよう、番人が「ここの語がどの op の照合語彙にも
#:   出てこないこと」を機械で確かめる（tests/test_we_say_so_when_we_cannot_write_it.py）。
#: ★★ その番人の**限界**（2026-09-09・語を足そうとして手で見つけた）:
#:   衝突試験が見るのは「**自分の語彙と重なるか**」だけで、
#:   「**別名で同じ能力を持っていないか**」は見えない。実例 ── 『フィルタ』は
#:   どの op の語彙にも無いので衝突なしと出るが、この道具は同じことを『抽出』で
#:   出来る。入れていたら正当な依頼を永久に断っていた。
#:   ★ だからここへ語を足す時は、機械の緑だけで通さず
#:     **「これを別の名前で出来ないか」を人が一度問う**こと。
#: ★ 文字色を**できるようにする**のは別の話（Namakoo 決裁 2026-09-09: 出荷後の
#:   アップデート枠）。能力を足しても、次の未対応属性で同じ事故が起きるため、
#:   まず「できないと言う」側を機械にする。
WE_DO_NOT_DO = {
    "文字色": "文字の色", "フォント色": "文字の色", "文字の色": "文字の色",
    "字の色": "文字の色", "フォントの色": "文字の色",
    "斜体": "斜体", "イタリック": "斜体",
    "下線": "下線", "アンダーライン": "下線",
    "取り消し線": "取り消し線", "打ち消し線": "取り消し線",
    "行の高さ": "行の高さ", "行高": "行の高さ",
    "文字サイズ": "文字の大きさ", "フォントサイズ": "文字の大きさ",
    "文字の大きさ": "文字の大きさ",
    "書体": "フォントの種類", "フォント名": "フォントの種類",
    # --- ここから下は属性でなく、表そのものでない設定（2026-09-09 追加）---
    # ★ 足す前に人が「別の名前で出来ないか」を問うた。落としたもの:
    #     フィルタ / オートフィルタ … 『抽出』で出来る（★ 二重語）
    #     グループ化               … 『集計』で出来る
    #     条件つき書式             … 『条件つき書換』と紛らわしい（単発では既に断れている）
    #     コメント・シート名・新しいシート … 列名や新シート作成と衝突しうる
    #   ★ 機械の衝突試験も 2 件落とした（『横向き』→「横」・『セルの結合』→「結合」）。
    "ウィンドウ枠": "ウィンドウ枠の固定", "ウインドウ枠": "ウィンドウ枠の固定",
    "枠の固定": "ウィンドウ枠の固定",
    "用紙の向き": "用紙の向き", "印刷範囲": "印刷範囲",
    "ハイパーリンク": "ハイパーリンク", "ズーム": "表示倍率",
    "シートの保護": "シートの保護",
    # ★ 2026-09-13（買い手役 2 回とも）: 「画像を挿入して」「ロゴを貼って」が 1 通目で
    #   「位置やサイズ／ファイルパスを教えてください」と**できると読める**断りになり、言われた
    #   とおり足すと 2 通目で「照合できません」── 往復が要らないよう、最初から名指しで断る。
    "画像": "画像の挿入", "写真": "画像の挿入", "ロゴ": "画像の挿入",
    "イラスト": "画像の挿入", "図形": "図形の挿入",
}


def asked_for_what_we_do_not_do(task: str, column_names=()) -> str | None:
    """依頼文が**この道具がやらないこと**を名指ししているなら、その名前を返す。

    ★ 列名に当たる語では鳴らさない（『書体』という列が実在する表なら、それは対象であって
      属性の指定ではない）── format_asked_but_not_supported と同じ作法に揃える。
    """
    text = task or ""
    cols = {str(c) for c in (column_names or ()) if c}
    for word, label in WE_DO_NOT_DO.items():
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


def op_effect_mismatch(task: str, ops, declaration: str,
                       vocab_by_op: dict, effects: dict) -> list:
    """依頼が名指しした操作と、実行した操作の**効果の種類**が食い違うなら、当たった語を返す。

    ★★ なぜ一般形が要るか（2026-09-07・効果の行列が本物の欠陥を 1 件掴んだ）:

        依頼   「机の行と棚の行を**交換して**」
        実行   操作:**行追加** 挿入位置:2 入れる値:品名=棚
        出力   成功（行が 4 → 5 に増えた）

      ★ 23 回中 22 回は正しく入れ替えており、**1 回だけ化けた**。取り除きだけを見る版
        （旧 removal_asked_but_not_done）では、入れ替えも行追加も「取り除かない」ので黙る。

    ★ 拒否は 3 つ（この 3 つで、実測の誤爆 24 件が全部消えた）:
      ① 実行した op **自身の語彙**が依頼文に在る（利用者がその操作を名指ししている）
      ② 2 つの op の**効果が重なる**（COMPUTE_COLUMN と ADD_COLUMN は共に列を作る＝上位下位）
      ③ 当たった語が**解釈行に出ている**（『小計の列を追加』の『小計』は新しい列の名前
         ＝依頼は宣言に反映されている ── 残差の関所と同じ考え）

    ★ 実測: 5,045 件の実走行で **21 件 0.42%・全部が本物**
      （DEDUP と EXTRACT が「消して」と言われて消さない 19 件 ＋ 入れ替えが行追加に化けた 2 件）。
    """
    text, decl = task or "", declaration or ""
    if not text:
        return []
    # ★★ 2026-09-07（今朝の凍結予測①が当たった）: **複合計画で誤爆した**。
    #   「単価で並べ替えして、数量と単価を掛けた金額の列を足して」は 2 段の計画で、
    #   依頼文が複数の操作を名指しするのは**当たり前**。1 段ぶんの op と比べると、
    #   もう一方の段の語が必ず食い違いに見える。
    #   ★ だから比べる相手は「その走行が**実行した op の集合**」にする。
    mine_ops = {str(o) for o in (ops if isinstance(ops, (list, tuple, set)) else [ops]) if o}
    if any(p for o in mine_ops for p in (vocab_by_op.get(o) or ()) if p and p in text):
        return []                                   # ① 自分の操作が名指しされている
    mine = set()
    for o in mine_ops:
        mine |= set(effects.get(o) or ())
    hits = []
    for other, phrases in (vocab_by_op or {}).items():
        theirs = set(effects.get(other) or ())
        if other in mine_ops or not theirs or (theirs & mine):
            continue                                # ② 効果が重なるなら食い違いでない
        hits += [p for p in (phrases or ()) if p and p in text and p not in decl]  # ③
    if not hits and mine and not (mine & {WRITE_REMOVE}):
        hits = [w for w in BARE_REMOVALS if w in text and w not in decl]
    return list(dict.fromkeys(hits))


#: 「取り除く」の効果名（登録簿の writes に入る値）。
WRITE_REMOVE = "remove"



# ─────────────────────────────────────────────────────────────────────
# 依頼に**無い**値が実体に書かれる向き（2026-09-10）
# ─────────────────────────────────────────────────────────────────────
#
# ★★ なぜ在るか（出荷前の実機テストが落ちて分かった・実測 11 回中 7 回）:
#
#     依頼   「5行目に丸山工業の行を作って」   ← 値の指定は**どこにも無い**
#     実物   5行目 = ['丸山工業', 1, 1000, 1000]
#     出力   **✓ 機械検証済み**
#
#   これまで塞いできたのは「依頼に在るものが宣言から落ちる」向きで、`residue` と
#   `op_effect_mismatch` がその隣を受け持っていた。**逆向き**は誰も見ていなかった。
#
# ★★ 実害は 2 つある（2 つ目の方が重い）:
#
#   ① 請求書に単価 1000・金額 1000 が入る。1×1000=1000 で**算数が合っている**ので、
#      行として自然に見え、人の目でも滑る。
#   ② ★ 捏造した値は `formula_columns_to_inherit` の除外集合に入るので、
#      **その列の式の継承が止まる**。実測（式を持つ表・6 回）:
#
#          捏造なし   金額 = '=B5*C5'   式が継承された
#          捏造あり   金額 = 1000       ★ 式が入らない（件数を直しても追随しない）
#
#      つまり捏造は、ゴミを足すだけでなく**表の整合そのものを殺す**。
#
# ★ 落とす向きの誤り（人が本当に指定した値を落とす）の方が怖いので、判定は**緩い側**に
#   倒す ── 漢数字・全角・桁区切りを開いてから照合し、当たらなかった値だけを落とす。
#   落とした列は**必ず名指しで人に見せる**（黙って空にするのは別の嘘になる）。
#
# ★ 数値は**部分文字列で照合しない**。`件数=1` が依頼文の `1000` に当たって
#   「接地した」と誤判定する（この判定を書いていて自分で踏んだ）。数として比べる。

_KANJI_DIGIT = {"〇": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
                 "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_KANJI_UNIT = {"十": 10, "百": 100, "千": 1000}
_KANJI_BIG = {"万": 10 ** 4, "億": 10 ** 8}
_KANJI_ALL = set(_KANJI_DIGIT) | set(_KANJI_UNIT) | set(_KANJI_BIG)


def _kanji_run_to_int(run: str):
    """漢数字の連なりを整数にする（「千」=1000・「二十」=20・「三万五千」=35000）。

    ★ 読めなければ None を返す（推測しない）。
    """
    total, section, digit, seen = 0, 0, None, False
    for ch in run:
        if ch in _KANJI_DIGIT:
            digit, seen = _KANJI_DIGIT[ch], True
        elif ch in _KANJI_UNIT:
            section += (digit if digit is not None else 1) * _KANJI_UNIT[ch]
            digit, seen = None, True
        elif ch in _KANJI_BIG:
            section += digit or 0
            total += (section or 1) * _KANJI_BIG[ch]
            section, digit, seen = 0, None, True
        else:
            return None
    if not seen:
        return None
    return total + section + (digit or 0)


def _numbers_in(text: str) -> set:
    """依頼文に現れる数を全部集める（半角化・桁区切り除去・漢数字を開いたうえで）。"""
    import re
    import unicodedata
    t = unicodedata.normalize("NFKC", text or "")
    t = t.replace(",", "").replace("，", "")
    out = set()
    for m in re.finditer(r"\d+(?:\.\d+)?", t):
        try:
            out.add(float(m.group(0)))
        except ValueError:
            pass
    # ★ 漢数字（「単価は千円で」→ 1000）。読めた連なりだけ足す。
    for m in re.finditer("[" + "".join(_KANJI_ALL) + "]+", t):
        v = _kanji_run_to_int(m.group(0))
        if v is not None:
            out.add(float(v))
    return out


def _as_number(value):
    """数として読めれば float、読めなければ None。"""
    import unicodedata
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = unicodedata.normalize("NFKC", str(value)).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def values_not_grounded_in_the_request(task: str, values) -> list:
    """依頼文に接地しない値の**列名**を返す（＝道具が発明した値）。

    Args:
        task: 人が書いた依頼文。
        values: 列名 → 入れる値 の対応（ADD_ROW の `values`）。

    Returns:
        接地しなかった列名のリスト（依頼文の順序ではなく `values` の順）。

    ★ 判定は緩い側に倒す ── 当たらなかったものだけを挙げる。
    """
    import unicodedata
    if not values or not task:
        return []
    hay = unicodedata.normalize("NFKC", task).replace(",", "").replace("，", "")
    nums = _numbers_in(task)
    out = []
    for col, v in values.items():
        if v is None or str(v).strip() == "":
            continue                     # 空は書かれないので見ない
        n = _as_number(v)
        if n is not None:
            # ★ 数は数として比べる（部分文字列だと 1 が 1000 に当たる）
            if not any(abs(n - x) < 1e-9 for x in nums):
                out.append(str(col))
            continue
        s = unicodedata.normalize("NFKC", str(v)).strip()
        if s and s not in hay:
            out.append(str(col))
    return out


# ── ★★ 2026-09-16: 「印を付けて」が**行追加**に化けて架空の行が入った（盲検の買い手役 2 体目）──
#
# 実測（買い手の台帳・再現 2/2・EXIT=0 で通った）:
#     「納期が2026/09/30より後の行のチェック列に★を入れて」
#       → 解釈: 操作:行追加 挿入位置:6 入れる値:チェック=★
#       → 受注台帳の 5 行目と 6 行目の間に、**取引先も品目も金額も空の行**が入った
#   ★ 買い手の言葉:「印を付けてと頼んだだけです」「月末の受注台帳は、間違えたら請求が狂って、
#     取引先に謝りに行く紙です」。値段の回答は **0 円**、理由はこの 1 件だった。
#
# ★ なぜ止まらなかったか: 「その op にしてよいか」を**依頼文に問い返す**器官（OP_META の
#   `requires_word`）は在るのに、宣言していたのは **PIVOT ただ 1 つ**だった。
#   行を足す op は「足す意図」を何も要求していない ── 器官が在って配線が無い形。
#
# ★ 線の引き方（実物で測って決めた・思いつきで足さない）:
#   repo の正当な行追加の検体 8 件に「〜列に／〜列を」は **1 件も無い**。
#   行を足す依頼は「新しい行に入れる値」を言うのであって、**書き込む先の列を名指ししない**。
#   逆に買い手の 2 通りはどちらも列を名指ししている（「チェック列に」「チェック列を」）。
#   ★ 比較の語（「金額が50000以上の…」）も同じ形の合図 ── こちらは SET_CELL_VALUE が
#     既に使っている線をそろえる（片方だけ守らない）。
#: ★★ 2026-09-16 実測で引き直した線。事故の形は「**行き先の列**（〜列に／〜列へ）＋
#:   **書き込む動詞**（入れて・付けて・書いて・入力して…）」── 表の構造は変えず、
#:   条件に合う行のあるセルに値を置く依頼。これが構造 op に化けると、頼んでいない
#:   行や列が増減する（買い手 2 体目の 0 円査定の理由がこれ 1 件だった）。
#: ★ 線は**助詞でなく動詞**の側に在る。初版は助詞だけ（列に/列を/列の/列は）で断って、
#:   正当な依頼 195 件のうち **56 件を誤って断った**（「備考という列を作って」「原価列を削除して」…）。
#:   書き込む動詞を必須にしたら、同じ 195 件で**誤り 0**・事故の形 6/6 を捕まえた。
#: ★ 助詞は に／へ／**を** まで取る ──「チェック列を★に書き換えて」も同じ事故で、
#:   検体に「言い直しても同じだった」と書いてあったのに初版の正規表現は拾えていなかった。
#:   「を」を取っても「備考列を削除して」は動詞が無いので素通りする（実測済み）。
#: ★ 動詞を必須にしたので op ごとの例外名簿が要らない ──
#:   「数量が100未満の行を削除して」（条件語あり・正当な DELETE_ROWS）も素通りする。
#:   名簿を持つと、次に足す op で必ず同じ穴が開く（片配線で何度も踏んだ形）。
_COLUMN_DEST = re.compile(r"([^\s、。「」『』のはがをに]{1,12})列\s*[にへを]")
_WRITE_VERB = re.compile(r"(入れ|付け|つけ|書[きいく]|記入|入力|セット|立て|マーク|埋め)")


def names_a_column_to_write_into(task: str) -> str | None:
    """依頼文が「〜列に／〜列へ」と**書き込む先の列**を名指ししているか（その語を返す）。"""
    m = _COLUMN_DEST.search(task or "")
    return (m.group(1) + "列") if m else None


#: ★★ 2026-09-16 の掃き出し: 表の**構造**を変える op（行や列が増減する）を実体から数えたら
#:   **9 つ**在り、そのどれも「その操作をしてよいか」を依頼文に問い返していなかった。
#:   買い手の実文「…の行のチェック列に★を入れて」を 9 つ全部に通したら、**4 つが通った**:
#:       DELETE_COLUMN（チェック列が消える）／ INSERT_ROWS（空行）／
#:       APPEND_TOTAL（合計行）／ COMPUTE_COLUMN（式を書く）
#:   ★ 朝に直したのは ADD_ROW **1 つだけ**だった ── 1 件ずつ追うと、残り 8 つはそのまま残る。
#:   ★ だから関所は op ごとに置かず、**構造を変える op すべて**が通る 1 箇所に置く。
#: 値は (その op が表に何をするか, 本当にそれを頼む時の書き方) ── 断り文はここから作る。
#: ★ 例を op ごとに持つのは、列を消す話に行を足す例を出さないため（最初の版はそうなっていた）。
STRUCTURAL_EFFECT = {
    "ADD_ROW": ("行を足す", "5行目に『丸山工業』の行を追加して"),
    "INSERT_ROWS": ("空の行を挿し込む", "3行目に空行を2行挿入して"),
    "APPEND_TOTAL": ("合計の行を足す", "最終行に『金額』の合計行を追加して"),
    "ADD_COLUMN": ("列を足す", "『備考』という列を右端に追加して"),
    "COMPUTE_COLUMN": ("計算した値の列を作る", "『単価』×『数量』を『小計』列として追加して"),
    "LOOKUP_FILL": ("別の表から写した列を作る", "『商品コード』で商品マスタを引いて『商品名』列を埋めて"),
    "SPLIT_CELL": ("セルを分けて列を増やす", "『住所』列をスペースで分けて2列にして"),
    "DELETE_ROWS": ("行を消す", "『金額』が空の行を削除して"),
    "DELETE_COLUMN": ("列を消す", "『備考』の列を削除して"),
}


def refuse_structural_op_that_is_really_a_write(task: str, op: str, *,
                                                comparison_in_task: bool) -> str | None:
    """表の構造を変える op にしてよいか。だめなら**断り文**を返す（続行してよければ None）。

    ★ 測った結果: 陽性（事故の形）6/6 を捕まえ、repo の正当な依頼 195 件の**誤りは 0**。
    ★ 断り文には「ではどう書けば通るか」を必ず入れる ── 断るだけでは買い手の仕事は進まない。
    """
    pair = STRUCTURAL_EFFECT.get(op)
    if pair is None:
        return None
    effect, example = pair
    if not _WRITE_VERB.search(task or ""):
        return None
    col = names_a_column_to_write_into(task)
    if not col and not comparison_in_task:
        return None
    what = "『" + col + "』に値を入れる形" if col else "条件（〜以上・〜未満 など）に合う行に値を入れる形"
    return (
        "依頼文が" + what + "で書かれているので、これは**" + effect + "**依頼ではなく"
        "「条件に合う行の、ある列を書き換える」依頼だと読みました ── "
        "そのまま実行すると" + effect + "ことになり、頼んでいない変化が表に残ります。実行しません。"
        + chr(10) +
        "  → 印を付けるなら「金額が50000以上の行のチェック列に『★』を付けて」のように、"
        "条件と列と値を書いてください（値は「」で囲みます）" + chr(10) +
        "  → 本当に" + effect + "なら、その言葉で書いてください（例:「" + example + "」）")


def refuse_row_add_that_is_really_a_write(task: str, *, comparison_in_task: bool) -> str | None:
    """★ 旧名（ADD_ROW 専用）── 掃き出しで 9 op に広げたので新しい方へ委譲する。"""
    return refuse_structural_op_that_is_really_a_write(task, "ADD_ROW",
                                                       comparison_in_task=comparison_in_task)
