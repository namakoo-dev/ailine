"""split_people — 『担当者別に分けて配る』の器官（需要⑤・2026-09-12）。
   設計: docs/DESIGN-20260912-担当者別に分けて配る.md（D1〜D6）。

★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）。openpyxl も
  触らない ── 行の並びを**値として**受け取り、(1) 分ける (2) 疑う (3) 証明の材料
  を返すだけ。Excel の読み書き・関所・書き出しは呼び出し側（ailine.cmd_split）が持つ。

★★ この器官が**しない**こと（設計 D6 ── この製品は代わりに決めない）:
  ・併合しない: `山田` と `山田　` は別の冊。`form_read.norm` が同じになるだけの組は
    **名指しする**（D4）── 寄せて 1 冊にするのは「代わりに決める」ことだから
  ・推測しない: 担当者の列は人が**見出しの文字**で指す。当たる見出しが 0 個／2 個以上
    なら分けない（D1）── 表からは決まらないものを決めない
  ・カナ↔漢字・敬称を寄せない: `ヤマダ`＝`山田` は「同じ人だと決める」こと。名指しも
    しない（決められないものを指ささない・D4 の★）。検体で出たら**取り逃し**として残す
  ・複数担当（`山田/佐藤`）はどの冊にも入れない ── 行番号と表記で人に返す
  ・空欄の行はどの冊にも入れない ── 空欄は誤配より安い（D2）
"""
from __future__ import annotations

from dataclasses import dataclass, field

from ailine_core import form_read, total_row
from ailine_core.primitives import is_number as _is_number
from ailine_core.total_row import _is_blank_cell as _blank

#: 書き手の印（docProps/core.xml の dc:creator）。★ 読む側の判定は stack.KIND_SIGNATURES。
CREATOR_MARK = "ailine split"
#: 焼き込む条件（description）の種別。
KIND = "split"

#: 一覧の冊（`_検分.xlsx`）── 拡張子は付けない（filetypes の登録簿から付ける）。
REPORT_STEM = "_検分"
REPORT_SHEET = "検分"
REPORT_HEADERS = ("種類", "対象", "行数", "金額", "なぜこうなったか")

#: 複数担当の区切り。★ 語の列挙ではなく**閉じた符号の集合**（設計 D2 が名指しした形）。
#:   判定は `form_read.norm` を通した値で見る ── 見た目が同じ中黒（`•` 等）も
#:   norm が `・` へ寄せるので、同じ 1 つの規則で当たる。
SEPARATORS = ("/", "／", "・", "、", ",")

#: ファイル名に使えない文字（Windows の予約文字 + 全角空白）。★ シート内の値は触らない。
_UNSAFE_IN_NAME = ('/', '\\', ':', '*', '?', '"', '<', '>', '|', '　')
_NAME_REPLACEMENT = "_"


@dataclass(frozen=True)
class SplitPlan:
    """1 冊をどう分けるかの計画と、分けなかったものの名指し。

    parts:      {担当者の表記（**原本の文字そのまま**）: [行番号, ...]}（現れた順）
    part_amounts: {表記: 金額の和}（★ 原本から見た参考値 ── 証明の『部分』は
                 呼び出し側が**書いた出力を読み戻して**数える。D5）
    blank:      担当者の列が空の行（どの冊にも入れない・D2）
    multi:      [(行番号, 表記)] 複数担当（決めない・D2）
    excluded:   分けない行の行番号 ── ① 全部空の行 ② 合計語（total_row の 1 箇所が決める）が
                どこかに在る行（★ 担当者の列そのものにラベルが入る表でも） ③ 非空セルが
                1 つだけの行（備考・小見出し・金額だけの小計 ── 明細の形をしていない）
    excluded_notes / blank_notes: [(行番号, 何が見えたか)] ── 検分に出す名指し
    lookalike:  [[表記a, 表記b], ...] norm が同じになる別の文字（併合はしない・D4）
    refused:    分けなかった理由（担当者の見出しが 0 個 or 2 個以上・D1）。None なら分けた
    unparsed:   金額が文字（`10,000円`）で数えられなかった行（読み替えない・D5）
    total_rows: [(行番号, 金額)] 元の表の合計行で金額が読めたもの（★ 明細の和と突き合わせる材料）
    whole_rows / whole_amount: 証明の『全体』（原本から数えた分母・D5）
    """
    header_row: int
    by_header: str
    by_column: int | None = None
    amount_header: str | None = None
    amount_column: int | None = None
    parts: dict = field(default_factory=dict)
    part_amounts: dict = field(default_factory=dict)
    blank: list = field(default_factory=list)
    blank_notes: list = field(default_factory=list)
    multi: list = field(default_factory=list)
    excluded: list = field(default_factory=list)
    excluded_notes: list = field(default_factory=list)
    lookalike: list = field(default_factory=list)
    refused: str | None = None
    unparsed: list = field(default_factory=list)
    total_rows: list = field(default_factory=list)
    by_note: str = ""            #: 列の選び方を人に言う 1 行（完全一致で絞った時だけ）
    blank_amount: float = 0.0
    multi_amount: float = 0.0
    whole_rows: int = 0
    whole_amount: float = 0.0
    headers: list = field(default_factory=list)

    @property
    def amount_counted(self) -> bool:
        """金額を数えたか（`--amount` が 1 列に決まった時だけ真・D5）。"""
        return self.amount_column is not None


def matching_columns(headers: list, wanted) -> list:
    """`wanted`（人が打った見出しの文字）に**当たる**列の 1 起点の列番号すべて。

    ★ 「当たる」の線（設計 D1 の例をそのまま機械にした）: `form_read.norm` が同じか、
      **norm が部分として含まれる**こと ── だから `担当` は `担当` にも `営業担当` にも
      当たる。当たりが 2 つ以上なら、どちらで分けるかは表から決まらない（→ 分けない）。
    ★ 完全一致を優先して 1 つに絞る形には**しない**。絞れば「担当」と「営業担当」が
      同居する表を黙って片方で分けてしまう ── 設計 D1 が名指しした最悪の形。
    """
    want = form_read.norm(wanted)
    if not want:
        return []
    hits = []
    for i, h in enumerate(headers, start=1):
        got = form_read.norm(h)
        if got and want in got:
            hits.append(i)
    return hits


def safe_filenames(values: list, reserved=()) -> dict:
    """{担当者の表記: ファイル名（拡張子なし）}。★ シート内の値は 1 文字も変えない。

    ファイル名に使えない文字（`/ \\ : * ? " < > |` と全角空白）は `_` にする。
    別の表記が同じ名前になったら、機械的な連番で逃がす（`stack.own_output_headers` と
    同じ作法 ── 黙って 1 つの冊へ混ぜない）。`reserved` の名前も避ける（`_検分` 等）。
    """
    taken = {str(r) for r in reserved}
    out = {}
    for value in values:
        base = str(value)
        for ch in _UNSAFE_IN_NAME:
            base = base.replace(ch, _NAME_REPLACEMENT)
        base = base.strip() or _NAME_REPLACEMENT
        # ★ 2026-09-13（2 回目の買い手役・事務職）: `--by 元ファイル` で分けると値が `04月.xlsx` なので
        #   冊名が `04月.xlsx.xlsx` になった。拡張子で終わる値は、その拡張子を落として名前にする。
        if base.lower().endswith((".xlsx", ".xlsm", ".xls", ".pdf", ".csv")):
            base = base.rsplit(".", 1)[0] or _NAME_REPLACEMENT
        name, n = base, 2
        while name in taken:
            name = f"{base}_{n}"
            n += 1
        taken.add(name)
        out[value] = name
    return out


#: 書き方が**複数ある**法人格だけ（2026-09-12）。`合同会社` のように 1 通りしか書かない語は
#:   ゆれを作らないので入れない。`工房` `事務所` も入れない ── 剥がすと
#:   『あかね工房』と『あかね』が同じ会社に見えてしまう（別の会社でありうる）。
_CORP_FORMS = ("株式会社", "有限会社", "㈱", "㈲", "㍿",
               "(株)", "（株）", "(有)", "（有）")


def entity_core(value) -> str:
    """法人格の書き方を落とした「名前の芯」。★ 併合には使わない ── **名指しの鍵**。

    ★★ 2026-09-12 の実測（実表 17 行）: `㈱アルファ` と `株式会社アルファ`、`デルタ㈱` と
      `㈱デルタ`（前置と後置）、`イータ(株)` と `イータ株式会社`、`ラムダ㈱` と `ラムダ株式会社` ──
      **16 冊のうち 8 冊が本来 4 社**だったのに、`norm` が違うので組にならず名指しできなかった。
      日本の実務でいちばん多いゆれがこれ。
    ★ `norm` そのものに畳まない: あちらは `_looks_like_org`・宛先の除外・請求元の同定にも使われ、
      測っていない所まで振る舞いが動く。**ここ（名指しの鍵）だけ**に閉じる。
    ★ 前置/後置の両方が落ちるので、語順の違いも同じ芯になる。
    """
    t = form_read.norm(value)
    for form in _CORP_FORMS:
        t = t.replace(form_head := form, "")            # noqa: F841 ── 前置も後置も落とす
    return t


#: 敬称 ── 末尾に付くと同じ人が別人に分かれる語。
#: ★ `氏` は**入れない** ── `高氏` のような姓を `高` に潰して偽の組を作る（得より害）。
#: ★ 落とすのは**末尾だけ**（語の中の「様」を壊さない）。
HONORIFICS = ("御中", "様", "さま", "サマ", "さん", "殿")

#: 異体字 ── 同じ姓の別の字。★ **小さく明示の表**にする（正規化を丸ごと当てると
#:   関係ない字まで畳んで偽の組が増える）。畳むのは**照合のときだけ**で、原本の文字は触らない。
VARIANTS = {"髙": "高", "﨑": "崎", "澤": "沢", "邊": "辺", "邉": "辺", "嶋": "島",
            "齋": "斎", "齊": "斉", "冨": "富", "龍": "竜", "德": "徳", "淸": "清"}


def matching_core(value) -> str:
    """照合のためだけの形 ── 空白を畳み、異体字を畳み、末尾の敬称を落とす。

    ★ 配る先の鍵ではない（束ねる鍵は**原本の文字そのまま**・D2）。ここで作るのは
      「同じ人の別の書き方かもしれない」を**名指しする**ための鍵だけ。
    ★ 合成する順: 空白 → 異体字 → 敬称。1 本に畳んでおくと、
      `髙橋様` と `高橋` のように**両方違う**組も 1 つの鍵で拾える。
    """
    text = "".join(VARIANTS.get(ch, ch) for ch in form_read.norm(value))
    for _ in range(len(HONORIFICS)):
        for word in HONORIFICS:
            if len(text) > len(word) and text.endswith(word):
                text = text[: -len(word)]
                break
        else:
            break
    return text


def lookalike_pairs(parts: dict) -> list:
    """同じものの**別の表記**の組（併合はしない・名指しだけ・D4）。

    2 つの鍵で束ねる ── どちらで束ねても「別の表記」であることに変わりはない:
      ① `norm` が同じ（空白・中黒・合字の違い）── 例 `緑川 誠` と `緑川誠`
      ② 法人格の書き方を落とすと同じ（`entity_core`）── 例 `㈱アルファ` と `株式会社アルファ`
      ③ 異体字を畳み末尾の敬称を落とすと同じ（`matching_core`）── 例 `髙橋` と `高橋`・
        `佐藤様` と `佐藤`（2026-09-13 に検体で測って足した）

    戻り値: [[表記a, 表記b], ...]（現れた順）。3 つ以上同じ鍵に落ちたら総当たりの組を出す。
    ★ 名指しは**安全側**の行い ── 別の会社が偶然同じ芯になっても、出るのは ⚠ 1 行で、人が見て捨てられる。
      併合（同じ冊に入れる）は取り返しが付かないので**しない**。
    """
    pairs, seen = [], set()
    for key in (form_read.norm, entity_core, matching_core):
        groups: dict = {}
        for value in parts:
            groups.setdefault(key(value), []).append(value)
        for members in groups.values():
            if len(members) < 2:
                continue
            for i in range(len(members)):
                for j in range(i + 1, len(members)):
                    mark = frozenset((members[i], members[j]))
                    if mark in seen:
                        continue
                    seen.add(mark)
                    pairs.append([members[i], members[j]])
    return pairs


def proof_breaks(proof: dict) -> list:
    """証明（部分の和＝全体）が破れている等式の名指し。戻り値が空なら破れていない。

    ★ 呼び出し側は『部分』を**書いた出力を読み戻して**入れる（D5）── 自分の書き込みを
      自分の記憶で数えた数を入れたら、この検算は恒真になる。
    """
    broken = []
    rows = proof.get("rows") or {}
    if rows:
        got = (rows.get("parts", 0) + rows.get("blank", 0)
               + rows.get("excluded", 0) + rows.get("multi", 0))
        if got != rows.get("whole"):
            broken.append(f"行: 全体 {rows.get('whole')} ≠ 部分の和 {got}"
                          f"（配った {rows.get('parts')}／空欄 {rows.get('blank')}／"
                          f"分けない {rows.get('excluded')}／複数担当 {rows.get('multi')}）")
    amount = proof.get("amount") or {}
    if amount and amount.get("counted"):
        got_a = (float(amount.get("parts") or 0) + float(amount.get("blank") or 0)
                 + float(amount.get("multi") or 0))
        whole_a = float(amount.get("whole") or 0)
        if abs(got_a - whole_a) > total_row.TOLERANCE:
            broken.append(f"金額: 全体 {whole_a} ≠ 部分の和 {got_a}"
                          f"（配った {amount.get('parts')}／空欄 {amount.get('blank')}／"
                          f"複数担当 {amount.get('multi')}）")
    return broken


def _refusal(headers: list, wanted, hits: list, what: str, *, exact_available: bool = False) -> str:
    """分けない理由の文（★ 見た見出しを並べる・当たった見出しは両方名指しする・D1）。"""
    seen = "／".join(f"『{h}』" for h in headers if str(h).strip()) or "（見出しが読めません）"
    if not hits:
        return (f"{what}の見出し『{wanted}』に当たる列がありません。"
                f"見た見出し: {seen}")
    names = "／".join(f"『{headers[i - 1]}』" for i in hits)
    tail = ("（打った文字とそのまま一致する見出しが 1 つあります ── その列でよければ `--exact` を"
            "付けてください。列番号では指せません）" if exact_available else
            "（両方とも同じ文字なら、元の表で片方の見出しを変えてから ── 列番号では指せません）")
    return (f"{what}の見出し『{wanted}』に当たる列が {len(hits)} つあります（{names}）"
            f"── どちらで分けるかは表からは決まりません。見出しの文字で 1 つに指してください" + tail)


def plan_split(grid_rows, header_row: int, by_header, amount_header=None, *,
               exact: bool = False) -> SplitPlan:
    """1 冊の行の並びから「どう分けるか」と「何を疑うか」を決める（値は作らない）。

    grid_rows: [(行番号, [その行の値, ...]), ...]。行番号 `header_row` の要素は
      見出しの並び（呼び出し側が `multifile.read_row_headers` で読んだもの ── 結合見出しを
      左上の値で埋めた形）、それ以外はセルの値そのもの。
    ★ ここは純関数 ── ファイルも openpyxl も触らない（同じ入力なら同じ計画）。
    """
    rows = {int(r): list(vals or []) for r, vals in grid_rows}
    headers = list(rows.get(header_row) or [])
    hits = matching_columns(headers, by_header)
    by_note = ""
    literal = [i for i in hits if form_read.norm(headers[i - 1]) == form_read.norm(by_header)]
    if len(hits) > 1 and exact and len(literal) == 1:
        # ★★ 2026-09-13（2 回目の買い手役・会計）: 『担当者』と『営業担当者』が並ぶ表で `--by 担当者` が
        #   「2 つあります」で止まり、案内は「元の表で片方の見出しを変えてから」── 顧問先から受け取った
        #   表の見出しは変えられない。
        #   ★ 設計 D1（既定では完全一致で絞らない ── 『担当』が『営業担当』と同居する表を黙って片方で
        #     分けるのが最悪の形）は**そのまま**。人が `--exact` で「打った文字とそのまま一致する見出しを
        #     使う」と明示した時だけ絞り、何を採ったか・何も当たっていたかを 1 行で言う（黙らない）。
        others = "／".join(f"『{headers[i - 1]}』" for i in hits if i != literal[0])
        by_note = (f"『{headers[literal[0] - 1]}』の列で分けました（{others} も当たりますが、"
                   "--exact の指定どおり、打った文字とそのまま一致する方を採りました）")
        hits = literal
    if len(hits) != 1:
        return SplitPlan(header_row=header_row, by_header=str(by_header), headers=headers,
                         amount_header=amount_header,
                         refused=_refusal(headers, by_header, hits, "担当者",
                                          exact_available=(len(literal) == 1)))
    by_column = hits[0]

    amount_column = None
    if amount_header not in (None, ""):
        amount_hits = matching_columns(headers, amount_header)
        if len(amount_hits) != 1:
            # ★ 金額の列が決まらないなら分けない ── 「どちらの列の金額か」を
            #   こちらで選べば、証明の分母そのものを推測で作ることになる。
            return SplitPlan(header_row=header_row, by_header=str(by_header), headers=headers,
                             by_column=by_column, amount_header=amount_header,
                             refused=_refusal(headers, amount_header, amount_hits, "金額"))
        amount_column = amount_hits[0]

    parts: dict = {}
    part_amounts: dict = {}
    blank, blank_notes, multi = [], [], []
    excluded, excluded_notes, unparsed, total_rows = [], [], [], []
    whole_rows = 0
    whole_amount = 0.0
    blank_amount = 0.0
    multi_amount = 0.0
    for row_num in sorted(r for r in rows if r > header_row):
        values = rows[row_num]
        whole_rows += 1
        by_value = values[by_column - 1] if len(values) >= by_column else None
        amount_value = (values[amount_column - 1]
                        if amount_column and len(values) >= amount_column else None)
        counted = _is_number(amount_value)

        filled = [v for v in values if not _blank(v)]
        if not filled:
            excluded.append(row_num)
            excluded_notes.append((row_num, "全部空の行"))
            continue
        # ★★ 分けない行①。**合計行を誰かの冊に入れるのが最悪の混入**（D2）。
        #   語の判定は total_row 1 箇所からしか来ない（ここで語を書き写さない）。
        #   合計行と見るのは 2 つの形だけ:
        #     (a) 担当者の列が**空**で、行のどこかに合計のラベルが在る（ふつうの合計行）
        #     (b) 担当者の列**そのもの**が合計のラベル（担当者の列が 1 列目の表 ── 検体 S05 で
        #         『合計』という名前の人の冊が出来た）
        #   ★ 2026-09-12: 実装の初版は「語が行のどこかに在れば、担当者が埋まっていても配らない」に
        #     広げていたが、探針で `設計費／合計商事／山田／1000` の行が山田の冊から**消えた**。
        #     受け取った人は消えた行に気づけない ── 誤爆は「⚠ 1 個」では済まない。狭い方に戻す。
        word = (total_row.row_has_total_word([by_value]) if not _blank(by_value)
                else total_row.row_has_total_word(values))
        if word:
            excluded.append(row_num)
            excluded_notes.append((row_num, word))
            # ★★ 2026-09-13（3 回目の買い手役・会計の MISSING #3「事務所がいちばん欲しい 1 行」）:
            #   元の表の合計行は黙って除外するだけで、**その値が明細の和と合っているか**を言わなかった。
            #   10 月分の冊に 9 月の合計 111,880 が持ち越されていても何も出ない ── 静かに壊れる側。
            #   ここでは数を控えるだけ（比べるのは報告の側・器は 1 箇所）。
            if counted:
                total_rows.append((row_num, float(amount_value)))
            continue
        if len(filled) == 1:
            # ★ 分けない行②（D2『備考』・『担当者列が空で金額列だけが在る行』）: 明細の行は
            #   日付・顧客・金額のように**複数の**セルが埋まる。1 セルだけの行は備考・小見出し・
            #   金額だけの小計 ── 明細の形をしていないので配らない（名指しはする）。
            excluded.append(row_num)
            excluded_notes.append((row_num, f"この行に在るのは『{filled[0]}』だけです"))
            continue
        if _blank(by_value):
            blank.append(row_num)
            blank_notes.append((row_num, _first_text(values)))
            bucket, key = "blank", None
        elif any(sep in form_read.norm(str(by_value)) for sep in SEPARATORS):
            # ★ 複数担当 ── 決めない。どの冊にも入れず、名指しして人に返す（D2）。
            multi.append((row_num, str(by_value)))
            bucket, key = "multi", None
        else:
            key = str(by_value)
            parts.setdefault(key, []).append(row_num)
            bucket = "part"

        # ★ 金額（ここから下は「配る／空欄／複数担当」の行だけ ── 分けない行の金額は
        #   全体にも部分にも入れない。合計行を足せば分母が黙って倍になる・D5）。
        if amount_column is None:
            continue
        if not counted:
            if not _blank(amount_value):
                # ★ 文字の金額（`10,000円`）── 読み替えず、数えず、名指しする（D5）。
                unparsed.append(row_num)
            continue
        amount = float(amount_value)
        whole_amount += amount
        if bucket == "blank":
            blank_amount += amount
        elif bucket == "multi":
            multi_amount += amount
        else:
            part_amounts[key] = part_amounts.get(key, 0.0) + amount

    return SplitPlan(header_row=header_row, by_header=str(by_header), by_column=by_column,
                     amount_header=(amount_header or None), amount_column=amount_column,
                     headers=headers, parts=parts, part_amounts=part_amounts,
                     blank=blank, blank_notes=blank_notes, multi=multi,
                     excluded=excluded, excluded_notes=excluded_notes,
                     lookalike=lookalike_pairs(parts), unparsed=unparsed, by_note=by_note,
                     total_rows=total_rows,
                     blank_amount=blank_amount, multi_amount=multi_amount,
                     whole_rows=whole_rows, whole_amount=whole_amount)


#: 配った冊の末尾に足す合計行のラベル（★ 語そのものは `total_row` が持つ ── ここで作る語も
#: あちらの判定に当たる物でなければ、書いた側と読む側が食い違う）。
OWN_TOTAL_LABEL = "合計"


def own_total_row(out_headers: list, plan, value, row_nums: list, row_values: dict):
    """配った冊の末尾に足す合計行（足さない回は None）。

    ★★ 2026-09-14（買い手役・会計の MISSING #2）: 「配った冊に本人の合計が無い。8800 は
      `_検分.xlsx` にしかない。本人に渡す冊なら合計が欲しい」。
    ★ 足すのは `--amount` を渡した回だけ（金額の列が決まっている回だけ）。
    ★ 担当者の値そのものが合計の語の回（検体 S05『合計商事』の家系）は足さない ── 人の名前と
      合計行が同じ文字になり、受け取った人が見分けられない。
    ★ 出所（元ファイル・元行）は**空**にする ── この行は元の冊から来ていない。空にした行は
      `is_own_total` が見分け、事後条件と独立検算の**両方**が値を確かめる（穴を開けない）。
    """
    if plan.amount_column is None:
        return None
    if total_row.row_has_total_word([value]):
        return None
    total = 0.0
    for r in row_nums:
        v = (row_values[r][plan.amount_column - 1]
             if len(row_values[r]) >= plan.amount_column else None)
        if _is_number(v):
            total += float(v)
    row = [None] * len(out_headers)
    row[plan.by_column - 1] = OWN_TOTAL_LABEL
    row[plan.amount_column - 1] = total
    return row


def is_own_total(src_row_value, values: list) -> bool:
    """その行が「冊の合計行」か ── 出所（元行）が空で、`OWN_TOTAL_LABEL` **そのもの**が在る行。

    ★★ 判断はここ 1 箇所（書く側・事後条件・独立検算の 3 つが同じ関数を呼ぶ）。
      3 箇所に書き写すと、片方だけ直って「出所の無い行が黙って通る」穴になる（開発手法 §13）。
    ★★ 2026-09-14: 初版は `total_row.row_has_total_word`（合計**らしさ**の規則）で見ていて、
      独立検算がそれを呼んだ瞬間に番人が赤くした ── 検算器が書き手の規則を再現すると恒真に
      近づく（`test_the_verifier_does_not_reproduce_the_rules`）。物差しは**自分が書いた
      ラベルという定数**に変えた（`PROVENANCE_HEADERS` と同じ ── 共有するのは語彙だけ）。
    ★ 出所が空でもラベルと違えば**偽** ── 捏造した行は今までどおり破れになる。
    """
    if _is_number(src_row_value):
        return False
    return any(str(v).strip() == OWN_TOTAL_LABEL for v in values if v is not None)


def total_row_check(plan) -> str | None:
    """元の表の合計行と、明細の和（配った＋空欄＋複数担当）を突き合わせる 1 行。

    戻り値: 人に見せる 1 行（合っていても言う）／合計行が無い・金額を数えていないなら None。
    ★★ 2026-09-13（3 回目の買い手役・会計）: 「10 月分の合計行は 111,880（9 月の持ち越し）で和は 94,720。
      黙って除外されるだけで『合計行の値が合いません』とは言われない ── 会計事務所がいちばん欲しい 1 行」。
    ★ 破れ（exit 5）にはしない ── 分けた結果は正しく、狂っているのは**元の表**。名指しして人に返す。
    ★ 金額が文字で数えられなかった行が在る回は「合わない」を言わない（分母が欠けている・出ない
      ことを信号にしない）。
    """
    if not plan.total_rows or plan.amount_column is None:
        return None
    if plan.unparsed:
        rows = "／".join(str(r) for r, _v in plan.total_rows)
        return (f"（元の表の合計行 {rows} 行目とは突き合わせていません ── 金額が文字の行が"
                f"{len(plan.unparsed)} 行あって明細の和が欠けています）")
    total = sum(v for _r, v in plan.total_rows)
    rows = "／".join(str(r) for r, _v in plan.total_rows)
    if abs(total - plan.whole_amount) <= total_row.TOLERANCE:
        return f"（元の表の合計行 {rows} 行目（{total:,.0f}）と明細の和が一致）"
    return (f"⚠ 元の表の合計行 {rows} 行目（{total:,.0f}）が明細の和（{plan.whole_amount:,.0f}）と"
            f"合いません（差 {total - plan.whole_amount:,.0f}）── 元の表の合計が古いか、"
            "明細に足し漏れがあります。分けた冊は明細のとおりです")


def _first_text(values: list) -> str:
    """行の中で最初に見えた値（検分で「この行には何が書いてあるか」を言うため）。"""
    for v in values:
        if not _blank(v):
            return str(v)
    return ""


def report_rows(plan: SplitPlan, files: dict, proof: dict | None = None) -> list:
    """一覧の冊（`_検分`）の行（見出しは REPORT_HEADERS）。★ 分母と名指しだけ・値は作らない。

    proof は呼び出し側が**書いた出力を読み戻して**組んだ証明（D5）。渡された時だけ
    証明の行を並べる ── ここで和を計算し直したら、それは証明ではなく 2 度目の記憶。
    """
    out = []
    if proof:
        rows = proof.get("rows") or {}
        out.append(["証明（行）", "部分の和 ＝ 全体", rows.get("whole"), "",
                    f"配った {rows.get('parts')}＋空欄 {rows.get('blank')}＋"
                    f"複数担当 {rows.get('multi')}＋分けない行 {rows.get('excluded')}"
                    f" ＝ {rows.get('whole')}"
                    f"（部分は配った冊を開き直して数えました／"
                    f"{'閉じています' if proof.get('ok') else '閉じていません'}）"])
        amount = proof.get("amount") or {}
        if amount.get("counted"):
            out.append(["証明（金額）", "部分の和 ＝ 全体", "", amount.get("whole"),
                        f"配った {amount.get('parts')}＋空欄 {amount.get('blank')}＋"
                        f"複数担当 {amount.get('multi')} ＝ {amount.get('whole')}"
                        "（分けない行の金額は入れていません）"])
        else:
            out.append(["証明（金額）", "数えていません", "", "",
                        "金額の見出しが指されていないので、金額は数えていません"
                        "（`--amount <見出し>` を付けると行数と同じやり方で証明します）"])
        for line in proof.get("broken") or ():
            out.append(["証明が破れた", "", "", "", line])
    for value, row_nums in plan.parts.items():
        amount = plan.part_amounts.get(value)
        out.append(["配った", value, len(row_nums),
                    amount if amount is not None else "",
                    f"{files.get(value, value)} へ書きました"])
    for row_num, seen in plan.blank_notes:
        out.append(["空欄", f"{row_num} 行目", 1, "",
                    f"担当者の列が空です（この行に見えるのは『{seen}』）"
                    "── どの冊にも入れていません"])
    for row_num, value in plan.multi:
        out.append(["複数担当", f"{row_num} 行目", 1, "",
                    f"『{value}』── 区切りを含みます。どちらの冊に入れるかは"
                    "表からは決まらないので、どの冊にも入れていません"])
    for row_num, seen in plan.excluded_notes:
        out.append(["分けない行", f"{row_num} 行目", 1, "",
                    f"『{seen}』── 明細の行ではないので配っていません"])
    for pair in plan.lookalike:
        rows_a = len(plan.parts.get(pair[0], ()))
        rows_b = len(plan.parts.get(pair[1], ()))
        out.append(["表記ゆれ", f"『{pair[0]}』／『{pair[1]}』", rows_a + rows_b, "",
                    "空白や中黒を無視すると同じ文字になります（別の冊のままにしています"
                    "── 同じ人だと決めるのは人の仕事です）"])
    for row_num in plan.unparsed:
        out.append(["金額が文字", f"{row_num} 行目", 1, "",
                    "金額が数値でないので和に数えていません（読み替えていません）"])
    return out
