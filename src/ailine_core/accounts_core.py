# -*- coding: utf-8 -*-
"""accounts_core — 『経費の勘定科目を先例から引く』の器官（需要③・2026-09-13）。
   設計: docs/DESIGN-20260913-経費の勘定科目を先例から引く.md
   ★ §6 が §2 を置き換える ── 区分の写像は §6.1、独立の検算は §6.5 に従う。

★ ailine を import しない・openpyxl も触らない（純関数・I/O なし）。行の並びを**値として**
  受け取り、(1) 候補を出す行を決め (2) 先例を集め (3) 区分の材料を組み、(4) 何を決めなかったかを
  名指しするだけ。ファイルの読み・冊の書き出し・関所は呼び出し側（ailine.cmd_accounts と
  ailine_core/accounts_read.py）が持つ。

★★ この器官が**しない**こと（設計 §0・§6.2 ── この製品は代わりに決めない）:
  ・辞書を持たない: 「タクシーは旅費交通費」を製品が知っている形は、会社ごとの流儀と
    ズレた瞬間に**代わりに決める**ことになる。根拠は使い手の先例だけ
  ・部分一致・類似・語の推測をしない: `form_read.norm` で畳んだ**完全一致**だけ。
    ゆれは `split_people.lookalike_pairs` で**名指しする**（併合はしない）
  ・ソフトの自動判別をしない: 列名の別名集合 1 つで受け、当たらなければ見た見出しを並べて断る。
    2 列に当たっても断る（split の D1 と同じ線）
  ・日付を読まない: 原文の文字列をそのまま運ぶ（`date_compare` に和暦略記を足さない ──
    あちらの発火条件は「実物で出たら」で、合成検体では鳴っていない）。だから「最新」は
    **読んだ並びの最後**を指す（日付として比べていない・根拠文にそう書く）
  ・貸方の勘定科目（払い方）は鍵に入れない ── 入れていないことを検分に書く（§6.3）

★★ 区分は `field_record.grade_of` **だけ**が決める（`tests/test_a_grade_has_exactly_one_source.py`
  が AST で縛っている）。ここで区分の語を組み立てない ── 出口が自分で語を書くと、
  導出を直しても片方が古い判断のまま残る（片配線）。

★★ 「同じ鍵の過去 N 件」は独立した出所ではない（設計 §6 の致命 1）:
  同じ取引先の過去 12 件は **1 つの入力の 12 度刷り**であって裏ではない。だから出所は
  **鍵ごと**に 1 つ数える（`Evidence.at` は鍵の名前）。N 件は根拠文の材料に落とす。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from ailine_core import field_record, form_read, split_people
from ailine_core.match import normalize_key as _normalize_key

#: 先例を引く鍵（★ 4 本・norm で畳んだ完全一致。設計 §6.1）。
#:   ★ 貸方の勘定科目（払い方）は**入れない** ── 払い方が同じだけでは費用の科目は決まらない。
KEYS = ("借方取引先", "貸方取引先", "借方補助科目", "摘要")

#: 候補を出す対象の列（空なら候補を出す）。
DEBIT_ACCOUNT = "借方勘定科目"
#: 候補を出す行の条件のもう半分（在ることを要求する ── 複合仕訳の継続行を触らないため）。
DEBIT_AMOUNT = "借方金額"
#: 借方の税区分（★ 2026-09-14・会計役の MISSING「税区分の整合」）。
#: ★ **無ければ黙る**（列が無い冊で文面を出さない ── 出ないことを信号にしない、の裏返し:
#:   測っていないものを測ったように言わない）。値は 1 文字も変えない。
DEBIT_TAX = "借方税区分"
#: 日付の列（★ 読まない・原文のまま根拠文へ echo するだけ）。
DATE = "取引日"
#: 伝票の番号（★ 触らない行を人が原本で見つけるための手がかり）。
SLIP_NO = "取引No"

#: 列名の別名集合。★ ソフトの自動判別はしない ── **1 つの集合**で受けて、
#:   当たらなければ見た見出しを並べて断る（設計 §6.2）。
#: ★ freee は別名が数個増えるだけで、**測る対象には入れない**（一次資料が 403 で読めず、
#:   推測した列名で検体を書けば測るのは推測の再現になる・設計 §6.2）。
_MONEY_TEXT = re.compile(r"^-?\d+(?:\.\d+)?$")


def money_value(v):
    """金額の列の文字を数にする（★ CSV は全部が文字 ── `ailine csv` の検疫と同じ線）。

    ★★ 2026-09-13（買い手役の初見）: 候補の冊の `借方金額(円)` が `'13750'`（文字）で書かれ、
      Excel で足せず、その冊を `split` に掛けると「金額 0 ＝ 0＋0＋0」に ✓ が付いた。
      同じ道具の `ailine csv` は数にする ── 片配線だった。
    ★ 桁区切り・円記号は外す。数に見えなければ**そのまま**（読み替えない・D5）。
    ★ 金額の列にしか掛けない（取引No の先頭 0 を落とさない）。
    """
    if not isinstance(v, str):
        return v
    s = v.strip().replace(",", "").replace("，", "").lstrip("¥￥").rstrip("円").strip()
    if _MONEY_TEXT.fullmatch(s):
        return float(s) if "." in s else int(s)
    return v


COLUMN_ALIASES = {
    DEBIT_ACCOUNT: ("借方勘定科目",),
    DEBIT_AMOUNT: ("借方金額(円)", "借方金額（円）", "借方金額"),
    DATE: ("取引日", "取引日付", "日付"),
    SLIP_NO: ("取引No", "取引No.", "伝票No.", "伝票番号"),
    "借方取引先": ("借方取引先",),
    "貸方取引先": ("貸方取引先",),
    "借方補助科目": ("借方補助科目",),
    "摘要": ("摘要", "借方摘要"),
    DEBIT_TAX: ("借方税区分", "借方税率区分"),
}

#: 見出しが無くても受ける唯一の形（弥生インポート形式）。★ 列位置で受けるのはここだけ。
#:   受ける条件は「1 行目が 25 列」かつ「1 列目（識別フラグ）が空または 4 桁以内の数字」──
#:   それ以外は断る（設計 §6.2 の線 ＝ ソフトの自動判別をしない・形だけで受ける）。
#: ★★ 1 列目の条件は「4 桁の数字」から**測って広げた**（2026-09-13・実装で踏んだ）:
#:   実物の識別フラグは 4 桁（2000 系）だが、**凍結済みの検体は 1 列目が空**だった
#:   （実測: 25 列・非空の位置は 2/4/6/9/… で、位置の対応（4=取引日付・5=借方勘定科目・
#:   6=借方補助科目・9=借方金額・17=摘要）は設計どおり）。狭いまま出すと弥生の冊を
#:   全部断ることになる ── 「測って動かない」だけを理由に広げ、**形だけで受ける**線は保つ
#:   （1 列目に文字が入っている 25 列の冊は今も断る・番人あり）。
YAYOI_COLUMN_COUNT = 25
#: ★ 税区分は 8 列目（一次資料の並び: 識別フラグ1／伝票No2／決算3／取引日付4／借方勘定科目5／
#:   借方補助科目6／借方部門7／**借方税区分8**／借方金額9）。
YAYOI_POSITIONS = {DATE: 4, DEBIT_ACCOUNT: 5, "借方補助科目": 6, DEBIT_TAX: 8,
                   DEBIT_AMOUNT: 9, "摘要": 17}
#: ★ 弥生の 1 列目（識別フラグ）は一次資料どおり **4 桁の半角数字**だけを受ける。
#:   実装の途中で検体の 1 列目が空だったため `{0,4}` に広げかけたが、それは「検体の欠けに
#:   合わせて規則を書く」形 ── 読み手は狭いまま、検体を実物に寄せた（2026-09-13）。
_YAYOI_FIRST_CELL = re.compile(r"^[0-9]{4}$")

#: 無ければ何も決まらない列（★ ここが 0 列なら断る）。
REQUIRED_COLUMNS = (DEBIT_ACCOUNT, DEBIT_AMOUNT)

#: 出力の冊（1 冊・シート 2 枚）。
SHEET_NAME = "候補"
#: 候補のシートに足す列（今回の行の右へ並べる）。
OUTPUT_HEADERS = ("候補の科目", "区分", "根拠", "先例の番地")
REPORT_SHEET = "検分"
#: ★ 「行」は**元ファイルの物理行**（説明行を含む）── 候補シートの行とは説明行の分だけずれる
#:   （2026-09-13・買い手役・会計）。見出しでそう言う。
REPORT_HEADERS = ("種類", "元ファイルの行", "件数", "内容")

#: 検分の「種類」。★ `空欄の理由` だけが分母に入る（候補が出なかった**候補行**）──
#:   `触らない行` は借方が埋まっている／継続行なので、そもそも分母の外（設計 §6.5 の 2）。
BLANK_KIND = "空欄の理由"
UNTOUCHED_KIND = "触らない行"
LOOKALIKE_KIND = "表記ゆれ"
#: 人が付けた科目が、その鍵の先例と食い違う行（★ 候補ではなく**検査**）。
#: ★★ 2026-09-14（買い手役・会計が 2 回）: 「この道具に一番期待したのは**先月と違う科目を付けて
#:   いないか**の確認」「『先月と違う科目を付けた行』の名指しが出力に無い」。
#:   借方勘定科目が既に埋まった行は「触らない行」として番号だけ控えて終わっていた ── 先例と
#:   突き合わせていなかった。★ 候補を出す行では候補＝先例なので「違う」は原理的に出ない。
DIFFERS_KIND = "付けた科目が先例と違う"
#: 税区分が、その鍵の先例と食い違う行（★ 科目が合っていても納税額が変わる ── 静かに乗る側）。
TAX_KIND = "税区分が先例と違う"
NOTE_KIND = "鍵について"
#: ★★ 2026-09-14（買い手役・会計の 中「検分の『13』」「候補シートと 1 ずれ・見出しに書いていない」）:
#:   行番号が**元ファイルの物理行**（説明行を含む）だと分かっても、**もう一方のシートでどこを
#:   見ればいいか**が分からない。見出しの語だけでは足りないので、表の**先頭**で 1 行だけ言う
#:   （読む人の疑問が生まれる場所に置く）。★ 番号の意味は変えない ── 説明を足す。
HOW_TO_READ_KIND = "この表の見方"
HOW_TO_READ = ("『元ファイルの行』は元の CSV／Excel の行番号です（説明行や見出しも数えた"
               "物理行 ── 1 行目から数えてください）。『候補』シートでは『元行』の列に同じ"
               "番号が入っています（その列で絞り込むと、同じ行に行けます）")

#: 書き手の印（docProps/core.xml の dc:creator）。★ 読む側の判定は `stack.KIND_SIGNATURES`。
CREATOR_MARK = "ailine accounts"
#: 焼き込む条件（description）の種別。
KIND = "accounts"

#: 先例の番地の書き方（★ 機械可読 ── `verify_accounts` がこれを読んでセルを見に行く）。
CITATION_SEP = "／"
_CITATION_RE = re.compile(r"^(?P<key>[^=]+)=(?P<rest>.+)$")
_CITATION_AT = "!"

#: 見出しが無い列を人に見せるときの言い方（`3列目`）。
_COLUMN_WORD = "列目"

#: 貸方を鍵に入れていないことの 1 行（★ 入れていないことを**書く**・設計 §6.3）。
CREDIT_NOTE = ("払い方（貸方の勘定科目）は鍵に入れていません ── 現金・未払金・カードが"
               "同じだけでは費用の科目は決まらないからです。貸方は鍵に入れていません"
               "（貸方取引先だけは鍵に入れています）")
#: 掃き出しの限界（★ 機械では確かめられない側を先に書く・設計 §6.5 の末尾）。
#: ★★ 2026-09-14（買い手役 2 体・別の役が別の冊で「同じ文の繰り返し」「336 字」）:
#:   冊ごとの但し書きが**行ごと・鍵ごとに刷られて**いた（1 行に 2〜3 回）。但し書きは
#:   冊で 1 回 ── 行の根拠には「読んだ並びで最後の先例」という**語そのもの**を残す
#:   （「最新」と呼ばないための語・2026-09-13 の買い手の指摘で入れた所は動かさない）。
ORDER_NOTE = ("先例の『読んだ並びで最後』は最新という意味ではありません ── 日付順には"
              "並べていません（冊に書かれている順に読んでいます）")
SWEEP_LIMIT = ("見ていない口（別の帳簿・人の判断・今回より後の仕訳）は掃けていません "
               "── 裏が取れたかの最後の一歩は人の目です")
#: ★★ 2 本の鍵が**同じ 1 行**を指していたときに必ず添える 1 行（設計に無い所見・実装で踏んだ）。
#:   設計 §6.1 の表は「2 本以上の鍵が当たれば 確」と決めているのでその判断は動かさないが、
#:   `field_record` の docstring は「写し合いの 2 つを裏と数えるな」と書いている ──
#:   同じ 1 行を 2 つの鍵で読むのは、まさにその形（結合セルを 2 度数えたのと同じ）。
#:   **区分は契約どおりに出し、人には見えるようにする**（黙って強く名乗らない）。
#: ★ 内訳が割れた鍵が在る行に必ず添える 1 行（その鍵は出所に数えないが、口は掃けていない）。
SPLIT_KEY_CAP = ("内訳が割れた鍵があるので『裏が取れた』とは呼びません"
                 "（割れた鍵は出所に数えず、ほかの鍵で引いています）")

#: 「確」に届かなかった理由（★ 画面の目盛り合わせに使う ── 値そのものは変えない）。
CEILING_ONE_KEY = "鍵が 1 本しか当たらなかった"
CEILING_SAME_ROW = "2 本以上の鍵が同じ 1 行を指した"
CEILING_SPLIT_KEY = "内訳の割れた鍵が残っている"
#: ★ 画面と検分で同じ順に並べる（1 箇所）。
CEILING_ORDER = (CEILING_ONE_KEY, CEILING_SAME_ROW, CEILING_SPLIT_KEY)

SAME_ROW_CAVEAT = ("この 2 本以上の鍵は同じ 1 行を指しています ── 別々の裏ではなく、"
                   "同じ 1 件の 2 通りの読み方です（裏が取れたと呼べるかは人が見てください）")


def key_identity(value):
    """鍵の同一性。★ 畳むのは `form_read.norm`、型は `match.normalize_key` の線。

    ★ 数値 123 と文字 "123" は**別の鍵**（黙って型変換しない ── match の Q15 と同じ線）。
    ★ 空（全角空白 1 文字も空）は None ＝ 鍵にならない。
    """
    marker = _normalize_key(value)
    if marker is None:
        return None
    text = form_read.norm(value)
    return (marker[0], text) if text else None


def cell(values, column):
    """1 起点の列番号でセルを引く（列が無ければ None）。★ 列の欠けを例外にしない。"""
    if not column:
        return None
    return values[column - 1] if len(values) >= column else None


def _seen_headers(headers) -> str:
    """見た見出しを並べた 1 行（★ 断るときは見たものを見せる・split の `_refusal` と同じ線）。"""
    names = "／".join(f"『{str(h).strip()}』" for h in (headers or []) if str(h or "").strip())
    return names or "（見出しが読めません）"


def near_headers(headers, role: str) -> list:
    """`role` に**近い**見出し（含む／含まれる）。★ 「目の前に在るのに気づかない」を塞ぐ。

    ★★ 出所（盲検 4 体目 ⑤）: 買い手の冊は見出しが『勘定科目』で、道具は
      「『借方勘定科目』に当たる列がありません……見た見出し: …『勘定科目』…」と言った。
      **探している物の候補をその場に並べておきながら、近いとは言わなかった** ──
      人は「在るじゃないか」で止まる。
    ★ 近さは含有だけで見る（曖昧な類似度を持ち込まない）。決めるのは人。
    """
    want = form_read.norm(role)
    out = []
    for i, h in enumerate(headers or [], start=1):
        got = form_read.norm(h)
        if not got or got == want:
            continue
        if got in want or want in got:
            out.append((i, str(h).strip()))
    return out


def _refuse_columns(headers, role: str, hits: list) -> str:
    """列が決まらないときの断り文（0 列 / 2 列以上のどちらも名指しする）。

    ★★ 2026-09-20（⑤）: 0 列の時に**通る道**を示すようにした。旧版は別名を並べるだけで、
      別名が役割名と同じ役割（借方勘定科目）では情報がゼロだった。
    """
    if not hits:
        near = near_headers(headers, role)
        lines = [f"『{role}』に当たる列がありません。見た見出し: {_seen_headers(headers)}"]
        if near:
            names = "／".join(f"『{n}』" for _i, n in near)
            lines.append(f"　★ 名前が近い列が在ります: {names}")
            lines.append(f"　→ その列でよければ `--column {role}={near[0][1]}` を付けて"
                         "もう一度実行してください")
        else:
            lines.append(f"　→ 読ませたい列があるなら `--column {role}=<見出し>` で教えてください"
                         f"（別名として自動で当たるのは: {'／'.join(COLUMN_ALIASES[role])}）")
        return "".join(lines)
    names = "／".join(f"『{str(headers[i - 1]).strip()}』" for i in hits)
    return (f"『{role}』に当たる列が {len(hits)} つあります（{names}）"
            f"── どちらを読むかは表からは決まりません。見出しを 1 つにしてください")


def match_column(headers, role: str) -> list:
    """`role` の別名に**完全一致**（norm 後）する列の 1 起点の列番号すべて。

    ★ 部分一致にしない: `借方金額` を部分一致で当てると `貸方金額` にも当たる列名が実在する。
      別名を並べる側で表記のゆれを吸う（`借方金額(円)` は別名として持つ）。
    """
    wanted = {form_read.norm(a) for a in COLUMN_ALIASES[role]}
    return [i for i, h in enumerate(headers or [], start=1)
            if form_read.norm(h) in wanted and form_read.norm(h)]


def parse_column_overrides(pairs) -> tuple:
    """`--column 役割=見出し` の並びを `{役割: 見出し}` にする。戻り値 (組, 断り)。

    ★ 役割の顔ぶれは `COLUMN_ALIASES` が唯一の出どころ（手で並べない ── 足した役割が
      黙って指定できないまま残らない）。

    ★★ 2026-09-22（盲検 6 体目 ⑥・致命）: 同じ役割を 2 回教えると**後勝ち**で上書きされ、
      「今回の冊＝『科目』／過去の冊＝『借方科目』」のように**出所が違う 2 種**を
      突き合わせる時に、どちらを教えても片方が落ちた。
      買い手の言葉:「うちの精算書と会計ソフト出力の見出しが一致することはまずありません」。
      ★ だから役割ごとに**候補の並び**を持つ（冊ごとに当たった方を使う）。
    """
    out: dict = {}
    for raw in list(pairs or ()):
        if "=" not in str(raw):
            return {}, (f"`--column` の形が違います（{raw}）── `役割=見出し` で書いてください"
                        f"（役割: {'／'.join(COLUMN_ALIASES)}）")
        role, name = str(raw).split("=", 1)
        role, name = role.strip(), name.strip()
        if role not in COLUMN_ALIASES:
            return {}, (f"『{role}』という役割はありません"
                        f"（役割: {'／'.join(COLUMN_ALIASES)}）")
        if not name:
            return {}, f"『{role}』に渡す見出しが空です（`--column {role}=<見出し>`）"
        out.setdefault(role, [])
        if name not in out[role]:
            out[role].append(name)
    return out, None


def resolve_accounts_columns(rows, overrides=None, notes=None) -> tuple:
    """行の並びから (見出し行, 見出しの並び, {役割: 列番号}, 断り) を決める。

    rows: [(行番号, [値, ...]), ...]（原本の 1 起点の行番号）。
    ★ 見出し行は「2 つ以上埋まった最初の行」── 弥生（見出し行が無い）だけは列位置で受ける。
    ★ 断ったら header_map は空（推測で先へ進まない）。
    """
    notes = notes if notes is not None else []
    ordered = sorted(rows, key=lambda rv: rv[0])
    if not ordered:
        return None, [], {}, "行がありません（空のファイル）"
    # ★ 2026-09-22: 見出し行と見出し名は**対で持つ**。片方だけ代入される経路が
    #   後から生えると、静かに空の見出しで先へ進む ── 型検査器がその予約を指した。
    head_row, headers = None, []
    for row_num, values in ordered:
        if sum(1 for v in values if form_read.norm(v)) >= 2:
            head_row = row_num
            headers = [v for v in values]
            break
    if head_row is not None:
        header_map, refusal = {}, None
        # ★★ 2026-09-20（⑤）: 人が `--column 役割=見出し` で教えた分は**そちらを採る**。
        # ★★ 2026-09-22（盲検 6 体目 ⑥・致命）: ただし**当たらない冊では自動照合へ落ちる**。
        #   ★ 事故: 道具が「`--column 借方勘定科目=科目` を付けて」と案内し、その通りに
        #     打つと**過去の仕訳の側が全部落ちた**（今回=自社の精算書『科目』／
        #     過去=会計ソフト出力『借方科目』）。★ 導線が嘘になっていた。
        #   ★ 旧版のコメントは「実際の冊は同じ書き出しなので、見出しも同じ」と書いて
        #     いたが、**出所が違う 2 種を突き合わせる道具**なので、揃う方が珍しい。
        #     買い手の言葉:「うちの精算書と会計ソフト出力の見出しが一致することは
        #     まずありません」。
        #   ★ 黙って別の列を読まない ── 落ちたことは notes で名指しして返す。
        for role, wants in (overrides or {}).items():
            wants = wants if isinstance(wants, (list, tuple)) else [wants]
            landed = None
            for want in wants:
                got = [i for i, h in enumerate(headers or [], start=1)
                       if form_read.norm(h) == form_read.norm(want)]
                if len(got) == 1:
                    header_map[role] = got[0]
                    landed = want
                    break
                if len(got) > 1:
                    return head_row, headers, {}, (
                        f"`--column {role}={want}` と教わりましたが、同じ見出しが"
                        f"{len(got)} 列あります ── どれを読むかは表からは決まりません")
            if landed is None:
                notes.append(
                    f"`--column {role}=" + "／".join(wants) + "` と教わりましたが、"
                    f"その見出しはこの冊にありません（この冊では自動で探しました）。"
                    f"見た見出し: {_seen_headers(headers)}")
        for role in COLUMN_ALIASES:
            if role in header_map:      # ★ 人が決めた分は照合し直さない
                continue
            hits = match_column(headers, role)
            if len(hits) == 1:
                header_map[role] = hits[0]
            elif len(hits) > 1:
                return head_row, headers, {}, _refuse_columns(headers, role, hits)
        missing = [r for r in REQUIRED_COLUMNS if r not in header_map]
        if not missing:
            return head_row, headers, header_map, None
        # ★★ 2026-09-22: 教えた見出しが**当たらず、しかも最後まで決まらなかった**役割は、
        #   note でなく**断りに載せる** ── 1 冊しか渡していない人にとって、当たらない
        #   指定は打ち間違いであり、黙って自動照合に落ちたら「教えたのに無視された」に
        #   見える（2026-09-20 に置いた契約: 在ると思って渡した人に、何が見えているかを返す）。
        #   ★ 2 冊以上なら当たった冊で使われるので、そちらは note のままでよい。
        stray = [n for n in notes if any(f"{r}=" in n for r in missing)]
        if stray:
            return head_row, headers, {}, "／".join(stray)
        # ★★ 2026-09-20（⑤）: 足りない列は**全部まとめて**言う。旧版は先頭 1 つだけを
        #   名指ししており、人は直して走らせて次を知る ── 往復が必要列の数だけ増える。
        #   ★ 照合の側は最初からそうしている（「決まらなかった役割を全部集めてから報告」）
        #     ── 同じ考えがこちらに配線されていなかった。
        refusal = "／".join(_refuse_columns(headers, r, []) for r in missing)
    else:
        refusal = f"見出しの行が見つかりません（最初の行: {_seen_headers(ordered[0][1])}）"

    # ★ 弥生（見出し行が無い・列位置固定 25 列）── **形だけ**で受ける。それ以外は断る。
    first_values = ordered[0][1]
    if len(first_values) == YAYOI_COLUMN_COUNT \
            and _YAYOI_FIRST_CELL.match(str(first_values[0] or "").strip()):
        return None, [], dict(YAYOI_POSITIONS), None
    # ★★ 2026-09-20（⑤）: 弥生の注記は**幅が弥生と同じ時だけ**出す。旧版は無条件に
    #   付けており、見出し行の在る 4 列の冊にまで「弥生の形でもありません」と言っていた ──
    #   買い手は弥生を使っていないので、**無関係な第二の診断**で誤導していた。
    if len(first_values) == YAYOI_COLUMN_COUNT:
        refusal += (f"（{YAYOI_COLUMN_COUNT} 列ですが、弥生の形としては受けられません"
                    "── 1 列目が 4 桁の数字である必要があります）")
    return head_row, (headers if head_row is not None else list(first_values)), {}, refusal


def column_labels(headers, header_map: dict, width: int, reserved=()) -> list:
    """候補のシートに出す列の名前（★ 見出しが無い冊は役割の名前と `N列目` で言う）。

    ★ 自分が足す列（`OUTPUT_HEADERS` と呼び出し側が渡す出所列）と衝突した原本の見出しは、
      **原本側**に機械的な番号を付けて逃がす ── 自分の列の名前が動くと、検算側が
      見つけられなくなる（`stack.own_output_headers` は逆に自分側をずらすが、こちらは
      検算がその名前で探すので動かせない）。
    """
    reserved = set(OUTPUT_HEADERS) | {str(r) for r in reserved}
    by_index = {i: role for role, i in header_map.items()}
    out, taken = [], set()
    for i in range(1, max(width, 0) + 1):
        raw = str((headers or [None] * width)[i - 1] or "").strip() if i <= len(headers or []) else ""
        name = raw or by_index.get(i) or f"{i}{_COLUMN_WORD}"
        if name in reserved or name in taken:
            name = f"{name}（{i}{_COLUMN_WORD}）"
        taken.add(name)
        out.append(name)
    return out


def format_citations(items) -> str:
    """先例の番地を 1 セルに畳む（★ 機械可読 ── `鍵=ファイル名!行`）。"""
    seen, parts = set(), []
    for key, name, row_num in items:
        mark = (str(key), str(name), int(row_num))
        if mark in seen:
            continue
        seen.add(mark)
        parts.append(f"{key}={name}{_CITATION_AT}{int(row_num)}")
    return CITATION_SEP.join(parts)


def parse_citations(text) -> list:
    """`format_citations` の逆（読めない断片は落とさず `None` の行番号で返す）。

    ★ 落として黙ると「番地が無い」が「番地が正しい」に化ける ── 出ないことは信号でない。
    """
    out = []
    for part in str(text or "").split(CITATION_SEP):
        part = part.strip()
        if not part:
            continue
        m = _CITATION_RE.match(part)
        if not m:
            out.append((part, None, None))
            continue
        rest = m.group("rest")
        if _CITATION_AT not in rest:
            out.append((m.group("key"), rest, None))
            continue
        name, _sep, row_text = rest.rpartition(_CITATION_AT)
        out.append((m.group("key"), name,
                    int(row_text) if row_text.strip().isdigit() else None))
    return out


@dataclass(frozen=True)
class AccountsPlan:
    """今回の冊をどう埋める候補があるかと、決めなかったものの名指し。

    records:   {行番号: field_record.Record}（★ 候補を出す行だけ ── 値と区分は Record が導く）
    citations: {行番号: [(鍵, 過去のファイル名, 元の行番号)]}（★ 機械可読の番地・§6.5 が読む）
    hits:      {行番号: 当たった先例の件数}（★ 根拠文の材料 ── 出所の数ではない）
    untouched: [(行番号, 理由)] 触らなかった行（埋まっている行・複合仕訳の継続行）
    lookalike: [(鍵, 表記a, 表記b)] norm/法人格/異体字で同じになる別の表記（併合はしない）
    notes:     検分に出す 1 行の並び（貸方を鍵に入れていないこと・掃き出しの限界）
    refused:   何もしなかった理由（None なら候補を出した）
    """
    header_map: dict = field(default_factory=dict)
    keys_used: tuple = ()
    records: dict = field(default_factory=dict)
    citations: dict = field(default_factory=dict)
    hits: dict = field(default_factory=dict)
    untouched: list = field(default_factory=list)
    differs: list = field(default_factory=list)     #: [(行番号, 1 行)] 付けた科目が先例と違う行
    tax_differs: list = field(default_factory=list)  #: [(行番号, 1 行)] 税区分が先例と違う行
    ceilings: dict = field(default_factory=dict)    #: {行番号: 確に届かなかった理由}
    lookalike: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    past_rows: int = 0
    past_precedents: int = 0
    refused: str | None = None

    @property
    def candidates(self) -> list:
        """候補を出す行の番号（★ 値が出た行だけではない ── 割 と 無 も候補行）。"""
        return sorted(self.records)


def _precedent_index(past_rows_by_file: dict, keys_used: tuple) -> tuple:
    """過去の行から {鍵: {同一性: [(科目の原文, ファイル名, 行番号, 日付の原文)]}} を組む。

    ★ 先例は「借方勘定科目が埋まっている行」だけ（★ 今回の冊の行はここへ渡さない ──
      自己先例は『確』の最短路。呼び出し側が同じファイルを両方に渡したら断る）。
    ★ 並びは読んだ順のまま ── 「最新」は**並びの最後**を指す（日付を読まないと決めたので）。
    """
    index = {k: {} for k in keys_used}
    rows_seen = filled = 0
    for name in past_rows_by_file:
        book = past_rows_by_file[name] or {}
        header_map = book.get("header_map") or {}
        for row_num, values in book.get("rows") or ():
            rows_seen += 1
            account = cell(values, header_map.get(DEBIT_ACCOUNT))
            if not form_read.norm(account):
                continue
            filled += 1
            date_text = str(cell(values, header_map.get(DATE)) or "").strip()
            # ★ 2026-09-14: 税区分を 5 番目に足す（既存の参照は位置で読んでいるので後方互換）。
            tax_text = str(cell(values, header_map.get(DEBIT_TAX)) or "").strip()
            for key in keys_used:
                ident = key_identity(cell(values, header_map.get(key)))
                if ident is None:
                    continue
                index[key].setdefault(ident, []).append(
                    (account, name, row_num, date_text or "（日付の列がありません）", tax_text))
    return index, rows_seen, filled


def _key_values(rows, column) -> dict:
    """その列に現れた表記（原本の文字そのまま・現れた順）。★ `lookalike_pairs` の入口。"""
    out = {}
    for _row_num, values in rows:
        raw = cell(values, column)
        if form_read.norm(raw):
            out.setdefault(str(raw), None)
    return out


def _why_untouched(values, header_map: dict) -> str:
    """触らない理由（★ 埋まっている行と、複合仕訳の継続行を**別の言葉**で言う）。"""
    account = cell(values, header_map.get(DEBIT_ACCOUNT))
    slip = str(cell(values, header_map.get(SLIP_NO)) or "").strip()
    tail = f"（{SLIP_NO} {slip}）" if slip else ""
    if form_read.norm(account):
        return f"{DEBIT_ACCOUNT}が既に『{str(account).strip()}』です{tail}"
    debit_cells = [header_map.get(r) for r in (DEBIT_ACCOUNT, DEBIT_AMOUNT, "借方補助科目",
                                               "借方取引先")]
    if not any(form_read.norm(cell(values, c)) for c in debit_cells if c):
        return (f"借方が丸ごと空の行です{tail} ── 複合仕訳の継続行は正当に空なので触りません")
    return f"{DEBIT_AMOUNT}が空です{tail} ── 金額の無い行に科目の候補は出しません"


def reason_of(record) -> str:
    """1 行の根拠（★ 区分の語は `field_record` から引く ── ここで語を作らない）。"""
    grade = field_record.grade(record)
    seen, parts = set(), []
    for evidence in record.evidences:
        if evidence.at in seen:
            continue
        seen.add(evidence.at)
        parts.append(evidence.how)
    if record.conflict_why:
        # ★★ 2026-09-13: `record.conflict` が真のときだけ出していたが、割れた鍵を**沈黙**
        #   させる形にしたら、値が出た行から「この支払先は割れています」が消えた（実測）。
        #   割れは出所に数えないだけで、**人に伝えるべき事実**は変わらない ── 常に出す。
        parts.append(record.conflict_why)
    if grade not in field_record.GRADES_WITH_VALUE and record.blank_reason:
        parts.append(record.blank_reason)
    if grade == field_record.CONFIRMED and record.swept_how:
        parts.append(record.swept_how)
    parts.append(f"区分 {grade}（{field_record.GRADE_MEANING[grade]}）")
    return "／".join(dict.fromkeys(parts))


def _amount_present(v) -> bool:
    """借方金額が「在る」か ── 数にして 0 でなければ在る。数に見えない文字（`10,000円`）も在る。"""
    m = money_value(v)
    if isinstance(m, (int, float)) and not isinstance(m, bool):
        return m != 0
    return bool(form_read.norm(v))


def plan_accounts(today_rows, header_map: dict, past_rows_by_file: dict) -> AccountsPlan:
    """今回の行（データ行だけ）と過去の行から、候補と名指しを決める（値は作らない）。

    today_rows:        [(行番号, [値, ...]), ...] ★ 見出し行を含めない（呼び出し側が外す）
    header_map:        {役割: 1 起点の列番号}（`resolve_accounts_columns` が決めたもの）
    past_rows_by_file: {ファイル名: {"header_map": {...}, "rows": [(行番号, [値, ...]), ...]}}
      ★ ファイルごとに列を解決する（同じソフトの書き出しでも列順が同じとは限らない）。
    ★ ここは純関数 ── ファイルも openpyxl も触らない（同じ入力なら同じ計画）。
    """
    keys_used = tuple(k for k in KEYS if header_map.get(k))
    absent = [k for k in KEYS if k not in keys_used]
    # ★ 2026-09-13（買い手役・会計・2 回とも）: 貸方取引先の列が無い弥生の冊でも定型文が
    #   「貸方取引先だけは鍵に入れています」と言い、次の行の「その列は無い」と**同じシートで食い違って**
    #   いた。文は見たものに合わせる（無い鍵のことは言わない）。
    credit_note = (CREDIT_NOTE if "貸方取引先" in keys_used
                   else CREDIT_NOTE.replace("（貸方取引先だけは鍵に入れています）",
                                            "（この冊には貸方取引先の列が無いので、貸方は鍵に入っていません）"))
    notes = [credit_note, ORDER_NOTE, SWEEP_LIMIT,
             ("鍵にした列: " + ("／".join(f"『{k}』" for k in keys_used) if keys_used
                                else "（1 本もありません）"))]
    if absent:
        notes.append("この冊に列が無いので引けなかった鍵: "
                     + "／".join(f"『{k}』" for k in absent)
                     + "（弥生の書き出しには取引先の列がありません ── 摘要と補助科目だけに"
                       "なるので、正直に 1 本の鍵しか無い行が増えます）")
    if not keys_used:
        return AccountsPlan(
            header_map=dict(header_map), notes=notes,
            refused=(f"先例を引く鍵の列が 1 本もありません（鍵: {'／'.join(KEYS)}）"
                     "── 列の取り違えか、別のソフトの書き出しが混ざっています"))

    index, past_rows, past_filled = _precedent_index(past_rows_by_file, keys_used)
    records, citations, hits, untouched, differs, tax_differs = {}, {}, {}, [], [], []
    ceilings = {}                       #: {行番号: 確に届かなかった理由}
    for row_num, values in sorted(today_rows, key=lambda rv: rv[0]):
        account = cell(values, header_map.get(DEBIT_ACCOUNT))
        amount = cell(values, header_map.get(DEBIT_AMOUNT))
        # ★ 候補を出すのは「借方勘定科目が空 かつ 借方金額が在る」行だけ（設計 §6.2）。
        #   継続行（借方が丸ごと空）と埋まっている行は触らず、番号と理由を控える。
        # ★★ 2026-09-13（3 回目の買い手役・会計）: CSV の継続行は借方金額が**文字の '0'** で、
        #   「在る」と数えて候補行にし、偽の「無」を作っていた（xlsx の数値 0 は触らない行）。
        #   同じ仕訳が CSV だと 13 行・xlsx だと 12 行 ── 数が入力の形に依っていた。金額の在否は
        #   `money_value` で数にしてから見る（0 は無い・数に見えない文字は在るとして隠さない）。
        # ★★ 税区分の検査は**候補を出す行と触らない行の両方**に効かせる（2026-09-14）──
        #   片方に配線して「直した」と書くのが この repo の再発する欠陥（開発手法 §13・§13c）。
        #   科目を出す／出さないに関わらず、納める税が変わる行は名指しする。
        tax_why = tax_differs_from_precedent(values, header_map, keys_used, index)
        if tax_why:
            tax_differs.append((row_num, tax_why))
        if form_read.norm(account) or not _amount_present(amount):
            untouched.append((row_num, _why_untouched(values, header_map)))
            # ★ 埋まっている行は「触らない」が、**先例と突き合わせる**（会計役が一番期待した所）。
            if form_read.norm(account):
                why = differs_from_precedent(values, header_map, keys_used, index, account)
                if why:
                    differs.append((row_num, why))
            continue
        record, cites, hit_count, ceiling = _record_for(values, header_map, keys_used, index)
        records[row_num] = record
        citations[row_num] = cites
        hits[row_num] = hit_count
        if ceiling:
            ceilings[row_num] = ceiling

    lookalike = []
    for key in keys_used:
        column = header_map[key]
        seen = _key_values(today_rows, column)
        for name in past_rows_by_file:
            book = past_rows_by_file[name] or {}
            past_column = (book.get("header_map") or {}).get(key)
            if past_column:
                seen.update(_key_values(book.get("rows") or (), past_column))
        # ★ 併合しない ── 名指しだけ（split と同じ器を**そのまま**呼ぶ・設計 §6.2）。
        for pair in split_people.lookalike_pairs(seen):
            lookalike.append((key, pair[0], pair[1]))

    plan = AccountsPlan(header_map=dict(header_map), keys_used=keys_used, records=records,
                        citations=citations, hits=hits, untouched=untouched, differs=differs,
                        tax_differs=tax_differs, ceilings=ceilings,
                        lookalike=lookalike, notes=notes, past_rows=past_rows,
                        past_precedents=past_filled)
    # ★★ 「当たった」は**先例の件数**で数える ── `evidences` で数えた初版は、鍵が当たって
    #   科目が割れた行（`conflict` は在るが `evidences` が空）を「1 件も当たらなかった」と
    #   読んで、割 しかない冊を丸ごと断っていた（実測・試験で捕まえた）。
    if records and not any(hits.get(row) for row in records):
        # ★★ 全行「無」の静かな成功をしない（設計 §6.3）── 文字コード違い・列の取り違え・
        #   別ソフトの混入がここに出る。断って人に返す。
        return AccountsPlan(
            header_map=plan.header_map, keys_used=keys_used, notes=notes,
            untouched=untouched, lookalike=lookalike, past_rows=past_rows,
            past_precedents=past_filled,
            refused=(f"どの行のどの鍵にも先例が 1 件も当たりませんでした"
                     f"（候補を出す行 {len(records)}／過去の行 {past_rows}・"
                     f"うち借方が埋まった行 {past_filled}／鍵 {'／'.join(keys_used)}）"
                     "── 文字コード違い・列の取り違え・別のソフトの書き出しの混入を"
                     "疑ってください"))
    return plan


def differs_from_precedent(values, header_map: dict, keys_used: tuple, index: dict,
                           account) -> str:
    """人が付けた科目が、鍵の先例と食い違うなら 1 行で返す（食い違わなければ空）。

    ★ 鳴らす条件は狭く取る ── その鍵の先例が**全部同じ科目**で、かつ付けた科目と違うときだけ。
      内訳が割れている鍵は**黙る**（「違う」と言えない ── 既存の「割れた鍵は沈黙」と同じ線）。
      先例が 1 件も無い鍵も黙る（新しい支払先）。
    ★ 値は 1 文字も変えない ── 触らない行のまま。決めるのは人（付け替えが正しい回もある）。
    """
    mine = form_read.norm(account)
    if not mine:
        return ""
    said = []
    for key in keys_used:
        raw = cell(values, header_map[key])
        ident = key_identity(raw)
        if ident is None:
            continue
        found = index[key].get(ident) or []
        if not found:
            continue
        by_account = {}
        for hit in found:
            by_account.setdefault(form_read.norm(hit[0]), []).append(hit)
        if len(by_account) != 1:
            continue                      # ★ 割れている鍵は沈黙
        theirs = next(iter(by_account.values()))
        if form_read.norm(theirs[0][0]) == mine:
            continue
        last = theirs[-1]
        said.append(f"{key}『{str(raw).strip()}』の先例 {len(theirs)} 件はすべて"
                    f"『{str(theirs[0][0]).strip()}』（{last[1]} {last[2]} 行目）")
    if not said:
        return ""
    return (f"付けた科目『{str(account).strip()}』と先例が違います: " + CITATION_SEP.join(said)
            + " ── 付け替えたのが正しいなら、このままで構いません（値は変えていません）")


def tax_differs_from_precedent(values, header_map: dict, keys_used: tuple, index: dict) -> str:
    """今回の行の税区分が、鍵の先例と食い違うなら 1 行で返す（食い違わなければ空）。

    ★★ 2026-09-14（買い手役・会計の MISSING「税区分の整合」）: 科目が合っていても税区分が
      違えば**消費税の納税額が変わる** ── しかも誰も気づかない（採用した人の申告に静かに乗る）。
    ★ 鳴らす条件は狭く取る（`differs_from_precedent` と同じ線）:
      その鍵の先例の税区分が**全部同じ**で、今回の行の税区分が**空でなく**違うときだけ。
      割れている鍵は沈黙（コンビニのように 10% と軽減 8% が本当に混ざる支払先が在る）。
    ★ 税区分の列が無い冊では何も言わない（`header_map` に無ければ `keys_used` を回る前に出る）。
    ★ 値は 1 文字も変えない ── 決めるのは人。
    """
    if not header_map.get(DEBIT_TAX):
        return ""
    mine = form_read.norm(cell(values, header_map[DEBIT_TAX]))
    if not mine:
        return ""                          # ★ 今回が空の行は黙る（空は誤値より安い）
    said = []
    for key in keys_used:
        raw = cell(values, header_map[key])
        ident = key_identity(raw)
        if ident is None:
            continue
        found = [hit for hit in (index[key].get(ident) or []) if len(hit) > 4 and hit[4]]
        if not found:
            continue
        by_tax = {}
        for hit in found:
            by_tax.setdefault(form_read.norm(hit[4]), []).append(hit)
        if len(by_tax) != 1:
            continue                       # ★ 割れている鍵は沈黙
        theirs = next(iter(by_tax.values()))
        if form_read.norm(theirs[0][4]) == mine:
            continue
        last = theirs[-1]
        said.append(f"{key}『{str(raw).strip()}』の先例 {len(theirs)} 件はすべて"
                    f"『{str(theirs[0][4]).strip()}』（{last[1]} {last[2]} 行目）")
    if not said:
        return ""
    return (f"税区分『{str(cell(values, header_map[DEBIT_TAX])).strip()}』と先例が違います: "
            + CITATION_SEP.join(said)
            + " ── 科目が合っていても納める税が変わります"
              "（値は変えていません・決めるのは人）")


def _record_for(values, header_map: dict, keys_used: tuple, index: dict) -> tuple:
    """1 行の記録を組む。戻り値 (Record, 番地の並び, 当たった先例の件数)。

    ★★ 出所は**鍵ごとに 1 つ**（`Evidence.at` は鍵の名前）── 同じ鍵の過去 N 件は
      1 つの入力の N 度刷りであって独立した裏ではない（設計 §6 の致命 1）。
    ★ 区分は `field_record.grade_of` だけが決める。ここでは材料（出所・食い違い）を渡すだけ。
    """
    evidences, cites, conflict_lines, looked, split_keys = [], [], [], [], []
    hit_count = 0
    for key in keys_used:
        raw = cell(values, header_map[key])
        ident = key_identity(raw)
        if ident is None:
            continue
        shown = str(raw).strip()
        found = index[key].get(ident) or []
        # ★ 2026-09-13（3 回目の買い手役・会計）: 「貸方取引先『アスクル』は過去に 1 件もありません」が
        #   帳簿にアスクルが 3 件在るのと矛盾して読めた（借方取引先の列には在る）。**列**の話だと言う。
        looked.append(f"{key}の列に『{shown}』{'は過去に 1 件もありません' if not found else ''}".rstrip())
        if not found:
            continue
        hit_count += len(found)
        by_account = {}
        for hit in found:
            by_account.setdefault(form_read.norm(hit[0]), []).append(hit)
        if len(by_account) == 1:
            same = next(iter(by_account.values()))
            last = same[-1]
            account = str(same[0][0]).strip()
            # ★ 「最新」とは**呼ばない**（2026-09-13・買い手役の初見）: 弥生は伝票 No. 順で出ることが
            #   あり、読んだ並びの最後が最新とは限らない。実際 `R08/08/25` が在るのに
            #   「最新 R08/07/25」と出た ── 同じ文で「日付は読んでいません」と開示していても、
            #   画面の主語が嘘なら嘘。「読んだ並びで最後の先例」と、そのままを言う。
            how = (f"{key}『{shown}』の先例は 1 件（{account}・{last[3]}・"
                   f"{last[1]} {last[2]} 行目）") if len(same) == 1 else (
                f"{key}『{shown}』は過去 {len(same)} 件すべて {account}"
                f"（読んだ並びで最後の先例 {last[3]}・{last[1]} {last[2]} 行目）")
            evidences.append(field_record.Evidence(rule=key, value=account, at=key, how=how))
            cites.append((key, last[1], last[2]))
        else:
            # ★ 1 本の鍵の内訳が 2 科目以上 → 割（`conflict=True`）。科目ごとの件数と
            #   最新の先例の日付を**原文のまま**添える（設計 §6.1 の★）。
            # ★★ 2026-09-13（実物の形で測って直した）: 内訳が割れた鍵は**沈黙**させる
            #   ── 出所に数えず、その行を止めもしない。理由に内訳だけ残す。
            #   実測: 実務の密度の検体（請求書 3 通から起こした仕訳）で到達 30%。落ちた分は
            #   「鍵が当たらない」ではなく「広い鍵（取引先）の割れが、完全に一致している
            #   狭い鍵（摘要）を道連れにしていた」── タクシー 3 行はすべて摘要が過去 1 行と
            #   完全一致して正しい科目を指していたのに空欄になっていた。
            #   ★ 実務では支払先が割れているのが普通（通販・タクシー・雑貨）。狭い鍵で解くのが
            #     作法で、そこを潰すと道具として使えない。
            #   ★ ただし**掃けていない口**なので、この行の区分は 単 を上限にする（下の cap）。
            #   ★ 本物の食い違い（2 本の鍵が違う科目を指す）は今までどおり割 ── そこは変えない。
            parts = []
            for same in by_account.values():
                last = same[-1]
                parts.append(f"{str(same[0][0]).strip()} {len(same)} 件"
                             f"（並びで最後の先例 {last[3]}・{last[1]} {last[2]} 行目）")
            split_keys.append(key)
            conflict_lines.append(f"{key}『{shown}』の内訳: " + "／".join(parts)
                                  + "（この鍵は出所に数えていません）")

    sources = {}
    for evidence in evidences:
        sources.setdefault(evidence.at, evidence.value)
    # ★★ 食い違い（割）になるのは 2 通り（2026-09-13・実物の形で測って直した）:
    #   ① 鍵どうしが違う科目を指した（本物の食い違い）
    #   ② **割れた鍵しか無い**（ほかの鍵が 1 本も引けていない）
    #   ★ ② を「無」にすると「先例がありません」という**嘘**になる ── 先例は在って割れている。
    #     凍結した検体の答えがここを守った（16 行が 割 → 無 に落ちて赤くなった実測）。
    #   ★ 一方、ほかの鍵が引けているなら割れた鍵は**沈黙**させるだけで行は止めない（上の枝）。
    conflict = len({str(v) for v in sources.values()}) > 1 or bool(split_keys and not evidences)
    same_row = len(sources) >= 2 and len({(name, row) for _k, name, row in cites}) == 1
    # ★★ 2026-09-14（買い手役・会計の 中「確 0/8（目盛り）」）: 一番上の区分が 1 行も出ない冊で
    #   「道具が何も見つけられなかった」と読める。実測すると 確 が稀なのは**規則が正しく
    #   働いた結果**（確＝2 本以上の鍵が**別々の先例**で一致・同じ 1 行の 2 度読みは裏 1 つ）。
    #   ★ だから確を出やすくする方へは触らない ── 直すのは**読み手の目盛り合わせ**。
    #   止まった理由はここ（決めている場所）で**構造として**控える。散文を後から読み直す形に
    #   すると、根拠の文言を変えた日に黙って壊れる。
    ceiling = ""
    if not conflict and sources:
        if same_row:
            ceiling = CEILING_SAME_ROW
        elif split_keys and len(evidences) > 1:
            ceiling = CEILING_SPLIT_KEY
        elif len(sources) == 1:
            ceiling = CEILING_ONE_KEY
    if same_row and not conflict:
        # ★★ 2 本以上の鍵が**同じ 1 行**を指していた ── 同じ行を 2 通りに読んだだけで、
        #   裏は 1 つ（`field_record` の「写しを裏に数えるな」と同じ形・結合セルの 2 度数え）。
        #   だから grade_of に渡す出所を 1 つに畳む → 区分は 単 に落ちる。
        #   鍵の名前は根拠に全部残す（何を見て 1 つと数えたかが人に見えるように）。
        first = evidences[0]
        merged_how = "／".join(e.how for e in evidences) + f"／{SAME_ROW_CAVEAT}"
        evidences = [field_record.Evidence(rule=first.rule, value=first.value,
                                           at=first.at, how=merged_how)]
    if split_keys and len(evidences) > 1:
        # ★ 割れた鍵が在る行は **単 を上限**にする ── 掃けていない口が残っているので
        #   「裏が取れた」とは名乗らない。出所を 1 つに畳んで grade_of に渡す。
        first = evidences[0]
        capped_how = "／".join(e.how for e in evidences) + f"／{SPLIT_KEY_CAP}"
        evidences = [field_record.Evidence(rule=first.rule, value=first.value,
                                           at=first.at, how=capped_how)]
    grade = field_record.grade_of(tuple(evidences), swept=True, conflict=conflict)
    blank_reason = ""
    if grade not in field_record.GRADES_WITH_VALUE:
        if len({str(v) for v in sources.values()}) > 1:
            blank_reason = ("鍵どうしが違う科目を指しています（"
                            + "／".join(f"{at}→{v}" for at, v in sources.items())
                            + "）── 値は出しません")
        elif split_keys:
            blank_reason = (f"先例の在る鍵（{'／'.join(split_keys)}）は内訳が割れていて、"
                            "ほかの鍵に先例がありません ── 値は出しません")
        elif looked:
            blank_reason = ("どの鍵にも先例がありません（見た鍵: "
                            + "／".join(looked) + "）")
        else:
            blank_reason = (f"鍵になる列（{'／'.join(keys_used)}）がすべて空の行です"
                            "── 引く手がかりがありません")
    # ★ 限界の文（`SWEEP_LIMIT`）は検分の注に 1 行在る ── 同じ文を 2 か所に置かない
    #   （行の根拠は「この行で何を見たか」だけにする・2026-09-14）。
    swept_how = (f"{len(keys_used)} 本の鍵（{'／'.join(keys_used)}）を見て、"
                 "食い違う鍵はありませんでした")
    record = field_record.Record(
        field=DEBIT_ACCOUNT, evidences=tuple(evidences), blank_reason=blank_reason,
        swept=True, swept_how=swept_how,
        conflict=conflict, conflict_why="／".join(conflict_lines))
    return record, cites, hit_count, ceiling


def confirmation_ceiling(plan: AccountsPlan) -> str:
    """「確」が 1 行も出なかった冊で、**なぜ 0 なのか**をその冊の数で言う 1 行（出なければ空）。

    ★★ 2026-09-14（買い手役・会計の「確 0/8（目盛り）」）: 0 が「道具が何も見つけられなかった」と
      読める。実測では 確 が稀なのは規則が正しく働いた結果（MF で 1/10・弥生で 0/6）なので、
      **確の条件は緩めない** ── 読み手の目盛りを合わせる（うま味調味料の最大化はしない）。
    ★ 数は `plan.ceilings`（決めた場所で控えた構造）から数える。散文は読み直さない。
    """
    # ★ 出す条件は「確が 0 行」だけ。★★ 初版は `or not plan.records` も書いていたが、
    #   変異試験で**素通り**した ── `ceilings` の鍵は候補行の部分集合なので、候補が 0 なら
    #   下の `not counts` で必ず空になる（到達できない枝だった）。番人を足すのでなく枝を消した
    #   （到達できない柵を「守っている」と名乗らない・2026-09-14 に 2 度目）。
    graded = [r for r in plan.records if field_record.grade(plan.records[r])
              == field_record.CONFIRMED]
    if graded:
        return ""
    counts = {}
    for reason in plan.ceilings.values():
        counts[reason] = counts.get(reason, 0) + 1
    if not counts:
        return ""
    breakdown = "／".join(f"{reason}行 {counts[reason]}"
                          for reason in CEILING_ORDER if counts.get(reason))
    # ★ 画面に出す文なので飾りの記号は使わない（`**` は文書の記法 ── 黒い画面では雑音）。
    return (f"確（2 本以上の鍵が別々の先例で一致）は 0 行 ── {breakdown}"
            f"（鍵は {len(plan.keys_used)} 本: "
            + "／".join(f"『{k}』" for k in plan.keys_used)
            + "）。この冊の形で決まる所で、見つからなかったという意味ではありません")


def candidate_rows(plan: AccountsPlan) -> list:
    """候補のシートに出す行の材料 ── [(行番号, 科目, 区分, 根拠, 番地)]。

    ★ 科目と区分は `field_record` が導いたものをそのまま運ぶ（出口で作り直さない）。
    """
    out = []
    for row_num in plan.candidates:
        record = plan.records[row_num]
        out.append((row_num, field_record.value(record), field_record.grade(record),
                    reason_of(record), format_citations(plan.citations.get(row_num) or ())))
    return out


def inspection_rows(plan: AccountsPlan) -> list:
    """検分のシートの行（見出しは `REPORT_HEADERS`）。★ 分母と名指しだけ・値は作らない。

    ★ `空欄の理由` の行番号が**分母の片側**（設計 §6.5 の 2）── 候補が出た行と合わせて
      「借方が空で金額が在る行」に一致するべきもの。検算はここから**名指しだけ**を読む。
    """
    out = [[HOW_TO_READ_KIND, "", "", HOW_TO_READ]]
    for row_num in plan.candidates:
        record = plan.records[row_num]
        if field_record.grade(record) in field_record.GRADES_WITH_VALUE:
            continue
        out.append([BLANK_KIND, row_num, plan.hits.get(row_num, 0), reason_of(record)])
    for row_num, why in plan.differs:
        out.append([DIFFERS_KIND, row_num, "", why])
    for row_num, why in plan.tax_differs:
        out.append([TAX_KIND, row_num, "", why])
    for row_num, why in plan.untouched:
        out.append([UNTOUCHED_KIND, row_num, "", why])
    for key, first, second in plan.lookalike:
        out.append([LOOKALIKE_KIND, "", "",
                    f"{key}: 『{first}』／『{second}』── 空白・法人格の書き方・異体字・"
                    "末尾の敬称を無視すると同じ文字になります（同じものだと決めるのは"
                    "人の仕事なので、別の鍵のままにしています ── 完全一致で引くので、"
                    "この 2 つは別の先例として数えています）"])
    ceiling = confirmation_ceiling(plan)
    if ceiling:
        # ★ 画面と冊に**同じ 1 行**（器は `confirmation_ceiling` 1 つ ── 2 度書かない）。
        out.append([NOTE_KIND, "", "", ceiling])
    for note in plan.notes:
        out.append([NOTE_KIND, "", "", note])
    return out


def grade_tally(plan: AccountsPlan) -> dict:
    """区分ごとの行数（★ 並びと語は `field_record.GRADE_ORDER` から ── ここで語を作らない）。"""
    counts = {g: 0 for g in field_record.GRADE_ORDER}
    for record in plan.records.values():
        counts[field_record.grade(record)] += 1
    return counts
