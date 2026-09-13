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
#: 日付の列（★ 読まない・原文のまま根拠文へ echo するだけ）。
DATE = "取引日"
#: 伝票の番号（★ 触らない行を人が原本で見つけるための手がかり）。
SLIP_NO = "取引No"

#: 列名の別名集合。★ ソフトの自動判別はしない ── **1 つの集合**で受けて、
#:   当たらなければ見た見出しを並べて断る（設計 §6.2）。
#: ★ freee は別名が数個増えるだけで、**測る対象には入れない**（一次資料が 403 で読めず、
#:   推測した列名で検体を書けば測るのは推測の再現になる・設計 §6.2）。
COLUMN_ALIASES = {
    DEBIT_ACCOUNT: ("借方勘定科目",),
    DEBIT_AMOUNT: ("借方金額(円)", "借方金額（円）", "借方金額"),
    DATE: ("取引日", "取引日付", "日付"),
    SLIP_NO: ("取引No", "取引No.", "伝票No.", "伝票番号"),
    "借方取引先": ("借方取引先",),
    "貸方取引先": ("貸方取引先",),
    "借方補助科目": ("借方補助科目",),
    "摘要": ("摘要", "借方摘要"),
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
YAYOI_POSITIONS = {DATE: 4, DEBIT_ACCOUNT: 5, "借方補助科目": 6, DEBIT_AMOUNT: 9, "摘要": 17}
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
REPORT_HEADERS = ("種類", "行", "件数", "内容")

#: 検分の「種類」。★ `空欄の理由` だけが分母に入る（候補が出なかった**候補行**）──
#:   `触らない行` は借方が埋まっている／継続行なので、そもそも分母の外（設計 §6.5 の 2）。
BLANK_KIND = "空欄の理由"
UNTOUCHED_KIND = "触らない行"
LOOKALIKE_KIND = "表記ゆれ"
NOTE_KIND = "鍵について"

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
SWEEP_LIMIT = ("見ていない口（別の帳簿・人の判断・今回より後の仕訳）は掃けていません "
               "── 裏が取れたかの最後の一歩は人の目です")
#: ★★ 2 本の鍵が**同じ 1 行**を指していたときに必ず添える 1 行（設計に無い所見・実装で踏んだ）。
#:   設計 §6.1 の表は「2 本以上の鍵が当たれば 確」と決めているのでその判断は動かさないが、
#:   `field_record` の docstring は「写し合いの 2 つを裏と数えるな」と書いている ──
#:   同じ 1 行を 2 つの鍵で読むのは、まさにその形（結合セルを 2 度数えたのと同じ）。
#:   **区分は契約どおりに出し、人には見えるようにする**（黙って強く名乗らない）。
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


def _refuse_columns(headers, role: str, hits: list) -> str:
    """列が決まらないときの断り文（0 列 / 2 列以上のどちらも名指しする）。"""
    if not hits:
        return (f"『{role}』に当たる列がありません（別名: "
                f"{'／'.join(COLUMN_ALIASES[role])}）。見た見出し: {_seen_headers(headers)}")
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


def resolve_accounts_columns(rows) -> tuple:
    """行の並びから (見出し行, 見出しの並び, {役割: 列番号}, 断り) を決める。

    rows: [(行番号, [値, ...]), ...]（原本の 1 起点の行番号）。
    ★ 見出し行は「2 つ以上埋まった最初の行」── 弥生（見出し行が無い）だけは列位置で受ける。
    ★ 断ったら header_map は空（推測で先へ進まない）。
    """
    ordered = sorted(rows, key=lambda rv: rv[0])
    if not ordered:
        return None, [], {}, "行がありません（空のファイル）"
    head_row = None
    for row_num, values in ordered:
        if sum(1 for v in values if form_read.norm(v)) >= 2:
            head_row = row_num
            headers = [v for v in values]
            break
    if head_row is not None:
        header_map, refusal = {}, None
        for role in COLUMN_ALIASES:
            hits = match_column(headers, role)
            if len(hits) == 1:
                header_map[role] = hits[0]
            elif len(hits) > 1:
                return head_row, headers, {}, _refuse_columns(headers, role, hits)
        missing = [r for r in REQUIRED_COLUMNS if r not in header_map]
        if not missing:
            return head_row, headers, header_map, None
        refusal = _refuse_columns(headers, missing[0], [])
    else:
        refusal = f"見出しの行が見つかりません（最初の行: {_seen_headers(ordered[0][1])}）"

    # ★ 弥生（見出し行が無い・列位置固定 25 列）── **形だけ**で受ける。それ以外は断る。
    first_values = ordered[0][1]
    if len(first_values) == YAYOI_COLUMN_COUNT \
            and _YAYOI_FIRST_CELL.match(str(first_values[0] or "").strip()):
        return None, [], dict(YAYOI_POSITIONS), None
    return head_row, (headers if head_row is not None else list(first_values)), {}, (
        refusal + f"（弥生の形（{YAYOI_COLUMN_COUNT} 列・1 列目が 4 桁の数字）でもありません）")


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
            for key in keys_used:
                ident = key_identity(cell(values, header_map.get(key)))
                if ident is None:
                    continue
                index[key].setdefault(ident, []).append(
                    (account, name, row_num, date_text or "（日付の列がありません）"))
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
    if record.conflict and record.conflict_why:
        parts.append(record.conflict_why)
    if grade not in field_record.GRADES_WITH_VALUE and record.blank_reason:
        parts.append(record.blank_reason)
    if grade == field_record.CONFIRMED and record.swept_how:
        parts.append(record.swept_how)
    parts.append(f"区分 {grade}（{field_record.GRADE_MEANING[grade]}）")
    return "／".join(dict.fromkeys(parts))


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
    notes = [CREDIT_NOTE, SWEEP_LIMIT,
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
    records, citations, hits, untouched = {}, {}, {}, []
    for row_num, values in sorted(today_rows, key=lambda rv: rv[0]):
        account = cell(values, header_map.get(DEBIT_ACCOUNT))
        amount = cell(values, header_map.get(DEBIT_AMOUNT))
        # ★ 候補を出すのは「借方勘定科目が空 かつ 借方金額が在る」行だけ（設計 §6.2）。
        #   継続行（借方が丸ごと空）と埋まっている行は触らず、番号と理由を控える。
        if form_read.norm(account) or not form_read.norm(amount):
            untouched.append((row_num, _why_untouched(values, header_map)))
            continue
        record, cites, hit_count = _record_for(values, header_map, keys_used, index)
        records[row_num] = record
        citations[row_num] = cites
        hits[row_num] = hit_count

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
                        citations=citations, hits=hits, untouched=untouched,
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


def _record_for(values, header_map: dict, keys_used: tuple, index: dict) -> tuple:
    """1 行の記録を組む。戻り値 (Record, 番地の並び, 当たった先例の件数)。

    ★★ 出所は**鍵ごとに 1 つ**（`Evidence.at` は鍵の名前）── 同じ鍵の過去 N 件は
      1 つの入力の N 度刷りであって独立した裏ではない（設計 §6 の致命 1）。
    ★ 区分は `field_record.grade_of` だけが決める。ここでは材料（出所・食い違い）を渡すだけ。
    """
    evidences, cites, conflict_lines, looked = [], [], [], []
    hit_count = 0
    for key in keys_used:
        raw = cell(values, header_map[key])
        ident = key_identity(raw)
        if ident is None:
            continue
        shown = str(raw).strip()
        found = index[key].get(ident) or []
        looked.append(f"{key}『{shown}』{'は過去に 1 件もありません' if not found else ''}".rstrip())
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
            # ★ 「最新」は**読んだ並びの最後**（日付は読まないと決めたので、日付として
            #   比べていない ── 根拠文にそう書く）。1 件のときは「すべて」と言わない。
            how = (f"{key}『{shown}』の先例は 1 件（{account}・{last[3]}・"
                   f"{last[1]} {last[2]} 行目）") if len(same) == 1 else (
                f"{key}『{shown}』は過去 {len(same)} 件すべて {account}"
                f"（最新 {last[3]}・{last[1]} {last[2]} 行目／"
                "最新は読んだ並びの最後です ── 日付は読んでいません）")
            evidences.append(field_record.Evidence(rule=key, value=account, at=key, how=how))
            cites.append((key, last[1], last[2]))
        else:
            # ★ 1 本の鍵の内訳が 2 科目以上 → 割（`conflict=True`）。科目ごとの件数と
            #   最新の先例の日付を**原文のまま**添える（設計 §6.1 の★）。
            parts = []
            for same in by_account.values():
                last = same[-1]
                parts.append(f"{str(same[0][0]).strip()} {len(same)} 件"
                             f"（最新 {last[3]}・{last[1]} {last[2]} 行目）")
                cites.append((key, last[1], last[2]))
            conflict_lines.append(f"{key}『{shown}』の内訳: " + "／".join(parts))

    conflict = bool(conflict_lines)
    sources = {}
    for evidence in evidences:
        sources.setdefault(evidence.at, evidence.value)
    same_row = len(sources) >= 2 and len({(name, row) for _k, name, row in cites}) == 1
    if same_row and not conflict:
        # ★★ 2 本以上の鍵が**同じ 1 行**を指していた ── 同じ行を 2 通りに読んだだけで、
        #   裏は 1 つ（`field_record` の「写しを裏に数えるな」と同じ形・結合セルの 2 度数え）。
        #   だから grade_of に渡す出所を 1 つに畳む → 区分は 単 に落ちる。
        #   鍵の名前は根拠に全部残す（何を見て 1 つと数えたかが人に見えるように）。
        first = evidences[0]
        merged_how = "／".join(e.how for e in evidences) + f"／{SAME_ROW_CAVEAT}"
        evidences = [field_record.Evidence(rule=first.rule, value=first.value,
                                           at=first.at, how=merged_how)]
    grade = field_record.grade_of(tuple(evidences), swept=True, conflict=conflict)
    blank_reason = ""
    if grade not in field_record.GRADES_WITH_VALUE:
        if conflict:
            blank_reason = "1 本の鍵の中で科目が割れています ── 値は出しません"
        elif len({str(v) for v in sources.values()}) > 1:
            blank_reason = ("鍵どうしが違う科目を指しています（"
                            + "／".join(f"{at}→{v}" for at, v in sources.items())
                            + "）── 値は出しません")
        elif looked:
            blank_reason = ("どの鍵にも先例がありません（見た鍵: "
                            + "／".join(looked) + "）")
        else:
            blank_reason = (f"鍵になる列（{'／'.join(keys_used)}）がすべて空の行です"
                            "── 引く手がかりがありません")
    swept_how = (f"{len(keys_used)} 本の鍵（{'／'.join(keys_used)}）を見て、"
                 f"食い違う鍵はありませんでした ── {SWEEP_LIMIT}")
    record = field_record.Record(
        field=DEBIT_ACCOUNT, evidences=tuple(evidences), blank_reason=blank_reason,
        swept=True, swept_how=swept_how,
        conflict=conflict, conflict_why="／".join(conflict_lines))
    return record, cites, hit_count


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
    out = []
    for row_num in plan.candidates:
        record = plan.records[row_num]
        if field_record.grade(record) in field_record.GRADES_WITH_VALUE:
            continue
        out.append([BLANK_KIND, row_num, plan.hits.get(row_num, 0), reason_of(record)])
    for row_num, why in plan.untouched:
        out.append([UNTOUCHED_KIND, row_num, "", why])
    for key, first, second in plan.lookalike:
        out.append([LOOKALIKE_KIND, "", "",
                    f"{key}: 『{first}』／『{second}』── 空白・法人格の書き方・異体字・"
                    "末尾の敬称を無視すると同じ文字になります（同じものだと決めるのは"
                    "人の仕事なので、別の鍵のままにしています ── 完全一致で引くので、"
                    "この 2 つは別の先例として数えています）"])
    for note in plan.notes:
        out.append([NOTE_KIND, "", "", note])
    return out


def grade_tally(plan: AccountsPlan) -> dict:
    """区分ごとの行数（★ 並びと語は `field_record.GRADE_ORDER` から ── ここで語を作らない）。"""
    counts = {g: 0 for g in field_record.GRADE_ORDER}
    for record in plan.records.values():
        counts[field_record.grade(record)] += 1
    return counts
