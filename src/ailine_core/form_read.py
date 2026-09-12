"""form_read — 帳票から項目を読む**規則**の層（2026-09-11）。

★★ 立ち位置:
      form_grid   … 座標と中身だけ。意味を持たない
      form_read   … ここ。「どこを見れば請求元か」を規則として持つ
      field_record… 集めた根拠から区分（確/単/割/無）を導く。**区分はそこだけ**

★★ 規則は**実物から**書いた（俺の頭の中の請求書からではない）。
  検体 v2 の「基礎」36 冊だけを見て書き、凍結してから残り 51 冊（敵対・税率・混入）で測る。
  ★ 見た分に当てはめた規則を全体の成績として報告しないため（~/.agents/skills/tuning-audit）。

★★ 実物が教えた、頭の中の請求書と違う点（すべて基礎 36 冊の実測）:

  ① **御中の付いたセルは会社名とは限らない。**
     `B3=ナギ商会株式会社` / `B4=経理部　御中` ── 御中が付くのは部署・担当者の側。
     ★ 2026-09-10 に「宛先 15/16 取れた」と誤報したのは、この御中セルを読んでいたから。
     ただし `B7=ナギ商会株式会社　御中` のように同じセルに同居する骨もある。**両方要る。**

  ② **住所ブロックは名前の上にも下にも来る。**
     misoca 系: 社名 → 〒 → 住所 → TEL   ／ inv21: 〒 → 住所 → 社名 → 登録番号
     → 「〒 の直上が請求元」は骨に依存した規則で、一般形ではない。

  ③ **宛先ブロックと請求元ブロックは同じ形をしている**（〒・住所・社名）。
     検体側が罠として明記している。だから難所は「探す」ではなく「**どちらがどちらか**」。

  ④ **ラベルのすぐ右が答えとは限らない。**
     `E38=消費税 | G38=0.1 | H38=6900` ── いちばん近い右は**税率**。
     `B12=ご請求金額 | (空) | (空) | E12=2200` ── 2 列空くこともある。
     → 「右へ進んで最初に出会う**数**」で当たるが、消費税の行だけは率を跨ぐ必要がある。

  ⑤ **ラベルの語中に空白が入る。** `O19=合 計 金 額` / `B11= ご請求金額　`（全角）
     / `B22=ご請求金額⏎（消費税込）`（改行）→ 照合前に**空白を全部落とす**。

  ⑥ **帯の値は同じ列に揃うとは限らない。**
     construction_bill: `C28=小計 E28=8000` / `G28=800`（列が違う）。

★★ ★ 検算のつもりが同じ計算になる罠（設計の要）:
  「小計 ＋ 消費税 ＝ 合計」を 2 つ目の根拠に数えたい。だが **合計セルが `=H37+H38`
  という式なら、一致するのは当たり前**で、検算ではない ── 2026-09-10 に踏んだ
  「自分の写しと一致して裏が取れた」と同じ形。
  → だから合計の**式**を読み、小計・消費税のセルを参照していたら
    その算術は `independent=False` として **2 つ目に数えない**。

★ ailine を import しない（ailine_core の作法）。
"""
from __future__ import annotations

import re
import unicodedata

from ailine_core.field_record import (GRADES_WITH_VALUE, SPLIT, Evidence,
                                      Record, grade_of)
from ailine_core.field_record import value as grade_value
from ailine_core.form_grid import Grid
from ailine_core.primitives import is_number as _is_number
from ailine_core.date_compare import parse_date_literal as _parse_date
from ailine_core.date_compare import parse_wareki_literal as _parse_wareki

# ── 語の正規化 ────────────────────────────────────────────
_SPACE = re.compile(r"[\s　 ]+")


#: 組織の形を表す**合字** ── `㈱`(U+3231) `㈲` `㍿`。NFKC が `(株)` `(有)` `株式会社` に
#:   ほどき、どれも `_CORP` に既に在る。★ 2026-09-12 の実測: ソフトが生成した実物の請求書が
#:   `㈱ミライ商事 御中` / `㈱テックコンサルティング` と書いていて、**宛先も請求元も取れなかった**。
#:   ★ 全体に NFKC をかけない（全角の数字が半角に化けて「金額の欄が文字」の拒否が消える）──
#:     中黒と同じく**閉じた文字クラス**だけをほどく。合字を含むセルは全群で 2 件（この 1 通のみ）。
_LIGATURES = re.compile(r"[㈠-㉃㍿]")

#: 見た目がほぼ同じ中黒（見出し `品番•品名` の `•` は U+2022 BULLET・実物に在る）。
#:   ★★ 2026-09-12 の実測: 器官の一覧は `品番・品名`（U+30FB）で、実物のベンダー雛形には
#:     `•`(U+2022) が 10 セル・検体 87 冊に 31 セル在った。**NFKC でも寄らない**
#:     （`･`→`・` は寄るのに `•` は寄らない）ため、明細の見出しが 85 冊中 42 冊しか
#:     見つかっておらず、**明細の掃き出し（いちばん強い番人）が半分の版面で黙っていた**。
#:   ★ これは言い回しの列挙ではなく**閉じた文字クラス**（同じ字を指す符号が複数ある）。
_DOTS = re.compile(r"[•·･‧∙⁃]")


def norm(s) -> str:
    """照合用に均す。★ 空白（半角・全角・改行）を**全部落とす**（実物の癖⑤）。

    ★ 中黒に見える符号は 1 つに寄せる（上の `_DOTS`）。組織の形の合字もほどく（`_LIGATURES`）。
    ★★ ここは**照合専用**で、人に出す値はここを通らない（`clean_org_name` は生の行から作る）。
      畳む前後で 87 冊の出す値が 1 つも変わらないことを実測して確かめた。
    """
    t = _LIGATURES.sub(lambda m: unicodedata.normalize("NFKC", m.group(0)),
                       _SPACE.sub("", str(s or "")))
    return _DOTS.sub("・", t)


#: 会社の形をした名前の語尾／語頭。★ 「これは組織の名前だ」の唯一の手がかり。
_CORP = ("株式会社", "有限会社", "合同会社", "合資会社", "合名会社",
         "(株)", "（株）", "(有)", "（有）", "医療法人", "社会福祉法人",
         "一般社団法人", "公益財団法人", "協同組合", "工房", "事務所")

#: 宛先を名指す印。★ これが付いたセル**そのもの**が名前とは限らない（癖①）。
_HONORIFIC = ("御中", "様")

#: 請求額の上部ラベル。
_LABEL_BILLED = ("ご請求金額", "請求金額", "御請求金額", "ご請求額", "今回ご請求額",
                 "今回請求額", "合計金額", "ご請求金額（消費税込）")
#: 小計のラベル。★ 「対象額（税抜）」を入れてはいけない ── あれは**列の見出し**で
#:   小計ではない。入れると税区分の 10% 欄を小計と誤認し、8% だけの冊で
#:   「明細 110,000 と小計 0 が合わない」と**偽の食い違い**を出す（実測 R01/R02）。
_LABEL_SUBTOTAL = ("小計", "小計金額", "税抜金額", "小計（税抜）")
_LABEL_TAX = ("消費税", "消費税額", "内消費税", "税額")
_LABEL_TOTAL = ("合計金額", "合計", "総合計", "税込合計", "ご請求金額")

#: 請求日・請求番号のラベル（2026-09-11・束で疑うために要る 4・5 つ目の項目）。
#:   ★ 事前測定: 請求日のラベルは検体 84/87・入れ子 14/14・実物の雛形 19/20 に在る。
_LABEL_DATE = ("請求日", "発行日", "請求年月日", "発行年月日", "日付")
_LABEL_NUMBER = ("請求番号", "請求書番号", "請求書No", "請求書No.", "請求No", "請求No.",
                 "伝票番号", "管理番号",
                 # ★★ 2026-09-12: 実物の請求書はラベルが**素の `No`** だけだった。
                 #   足すと明細の見出し（`No 品 目 数 量 …`）の右＝`品目` を拾う危険が在るが、
                 #   下の「値に数字を要求する」で弾ける（実測: 実物は 4 桁の数・雛形は `品目`）。
                 "No", "No.", "NO", "NO.", "№", "Ｎｏ", "Ｎｏ.")
#: 識別子らしさ ── 数字を 1 つも含まないものは請求番号ではない。
_HAS_DIGIT = re.compile(r"[0-9０-９]")

#: 雛形の埋め草として日付欄に残る形（`××年1月1日` など）。日付ではない。
_DATE_PLACEHOLDER_CHARS = "×〇○□■＊*"

#: 登録番号（インボイス制度）。★ 請求元の側にしか付かない ── 買い手には付かない。
_REGNO = re.compile(r"^T\d{13}$")


def _looks_like_org(text: str) -> bool:
    t = norm(text)
    return any(k in t for k in _CORP)


#: 名前の行ではないと分かる語。★ その行は連絡先・役割であって、社名そのものではない。
#: ★ 敬称（御中・様）はここに入れない ── 宛先の側では**敬称の付いた行こそが名前**。
#:   請求元の側で敬称つきを避けたいなら、呼ぶ前に弾く（read_issuer がそうしている）。
#:   ここに入れると、1 関数に畳んだときに片方の都合がもう片方を壊す（実測 13 冊）。
_NOT_A_NAME_LINE = ("登録番号", "担当", "TEL", "ＴＥＬ", "FAX", "ＦＡＸ", "E-mail",
                    "E-Mail", "〒", "電話", "取次", "代理店")

#: 雛形のプレースホルダに使われる字。★ これしか無い名前は名前ではない。
_PLACEHOLDER = "〇○●◯番地×✕＊*_－-― 　"


def name_lines(raw: str) -> list:
    """1 セルの中身を行に割り、**名前になり得る行だけ**返す。

    ★ 実物は 1 セルに改行で同居する（実測）:
        `高梨産業株式会社⏎(登録番号:T7010001234567)`
      初版はセルの中身をそのまま名前として返し、登録番号ごと持ち出した。
    """
    out = []
    for line in str(raw or "").splitlines():
        t = _strip_honorific(line).strip()   # ★ 敬称は行を落とす前に外す
        if not t:
            continue
        if _REGNO.match(norm(t)):
            continue
        if any(k in t for k in _NOT_A_NAME_LINE):
            continue
        out.append(t)
    return out


def is_placeholder(name: str) -> bool:
    """雛形のまま（`株式会社 〇〇〇`）か。★ 名前の部分が飾り字しか無い。"""
    body = norm(name)
    for k in _CORP:
        body = body.replace(norm(k), "")
    body = body.strip()
    if not body:
        return True
    return all(ch in _PLACEHOLDER for ch in body)


def clean_org_name(raw) -> tuple:
    """1 セルの中身から**組織の名前**を取り出す。★ 名前を作る道はここ 1 本だけ。

    ★★ なぜ 1 本に畳むか（2026-09-11、一度も測っていない実物の雛形 15 冊で踏んだ）:
      プレースホルダ（`○○株式会社`）の番人を **請求元の側にだけ**書いていた。
      宛先の側は同じ判定を持たず、空の雛形に対して「宛先＝○○株式会社」を
      **値として出していた**。
      ★ この repo は「二重化した経路は片配線が既定で起きる」を何度も踏んでいる。
        処方は**両方直すことではなく、1 関数に畳んで呼び出し側に選ばせないこと**。
        番人も 1 本の試験で両方の項目を縛る（片方を壊せば必ず赤になる）。

    戻り値: (名前, 使えない理由) ── 理由が空でなければ、その名前は使えない
    """
    # ★ 敬称を外すのは `name_lines` の中で 1 回だけ。ここで**もう一度**外すと
    #   守りが二重になり、片方を壊しても緑のままになる（番人が何も見なくなる）。
    #   実測: 敬称の扱いを 3 か所に重ねていて、どれを壊しても試験が通った。
    lines = [ln for ln in name_lines(raw) if _looks_like_org(ln)]
    if not lines:
        return "", "組織の名前の行が見つかりません（連絡先や役割の行だけです）"
    name = lines[0]
    if not name:
        return "", "敬称を除くと何も残りません"
    if is_placeholder(name):
        return "", f"雛形のまま（{name}）で、実際の社名ではありません"
    return name, ""


def _strip_honorific(text: str) -> str:
    t = str(text or "")
    for h in _HONORIFIC:
        t = t.replace(h, "")
    return t.strip().strip("　")


#: 金額の枠に入っていたら「答えの枠は埋まっている」とみなす文字の形。
#:   ★ `¥1,320,000-` / `１，３２０，０００`（全角）/ `#REF!` など。
_MONEYISH = re.compile(r"^[¥￥\$]?[\d０-９][\d０-９,，\.．\-ー－\s　]*[-ー－]?$")
_ERRORISH = ("#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!")


def _looks_like_money_text(v) -> bool:
    """文字で入った金額か、式の壊れか。★ 跨いではいけない中身。"""
    if not isinstance(v, str):
        return False
    t = norm(v)
    if not t:
        return False
    return t in _ERRORISH or bool(_MONEYISH.match(t))


def _first_number_right(grid: Grid, cell, span: int = 10, skip_rate: bool = True):
    """ラベルの右へ進んで最初の**数**（癖④）。

    ★ `skip_rate` は消費税の行だけの措置 ── `消費税 | 0.1 | 6900` の 0.1 は率であって
      額ではない。0 < v < 1 の数は率とみなして跨ぐ。

    ★★ 2026-09-11、未見 51 冊で踏んだ事故への処方:
      答えの枠に **文字で入った金額**（`¥1,320,000-` / 全角数字）や **`#REF!`** が
      在るとき、初版はそれを黙って跨いで**別のセルの数**を読み、しかも「確」と言った。
      → 枠が埋まっているなら、そこで**止まる**。読めないなら読めないと言う
        （黙って別のものを読むのが、この repo でいちばん高くつく失敗）。
    戻り値: (セル, 事情) ── 事情が空でなければ、値としては使えない。
    """
    for c in grid.right_of(cell, span=span):
        if _is_number(c.value):
            if skip_rate and 0 < float(c.value) < 1:
                continue                       # 税率は額ではない
            return c, ""
        if _looks_like_money_text(c.value):
            t = norm(c.value)
            if t in _ERRORISH:
                return c, f"{c.at} が {t} になっています（式が壊れています）"
            return c, f"{c.at} の金額が数値ではなく文字で入っています（{str(c.value)[:20]}）"
    return None, ""


def _labelled_number(grid: Grid, labels, *, rows=None, skip_rate=True) -> list:
    """ラベルに一致する文字セルを探し、その右の最初の数を返す。

    戻り値: [(ラベルのセル, 値のセル, 事情)] ── 事情が空でなければ値として使えない
    """
    out = []
    wanted = {norm(x) for x in labels}
    seen = set()
    for t in grid.text_cells():
        if rows and not (rows[0] <= t.row <= rows[1]):
            continue
        if norm(t.value) not in wanted or t.anchor in seen:
            continue
        seen.add(t.anchor)                     # ★ 結合の写しで同じラベルを 2 度数えない
        v, note = _first_number_right(grid, t, skip_rate=skip_rate)
        if v is not None:
            out.append((t, v, note))
    return out


def _labelled_text_right(grid: Grid, labels, *, rows=None, span: int = 4) -> list:
    """ラベルの**右**の最初の中身、または**同じセルにラベルと同居**している中身を返す。

    戻り値: [(ラベルのセル, 値のセル, 生の値)]
    ★★ 「下」には落ちない（2026-09-11 の事前測定）:
      請求日：／請求番号：はラベルが縦に積まれ、値は右に在る。右が空のとき下へ落ちると
      **次のラベル**や**発行者名**を値として拾う（probe で 77 件・58 件が全部それだった）。
      読めないなら読めないと言う方が安い。
    """
    out, seen = [], set()
    wanted = {norm(x).rstrip("：:") for x in labels}
    for t in grid.text_cells():
        if rows and not (rows[0] <= t.row <= rows[1]):
            continue
        if t.anchor in seen:
            continue
        raw = str(t.value)
        head = norm(raw).rstrip("：:")
        glued = None
        if head not in wanted:
            # 同居: 「請求日：2026/8/31」── ラベルで始まり、残りが在る
            for w in wanted:
                for sep in ("：", ":"):
                    pre = w + sep
                    if norm(raw).startswith(pre) and len(norm(raw)) > len(pre):
                        glued = raw.split(sep, 1)[1].strip()
                        break
                if glued is not None:
                    break
            if glued is None:
                continue
        seen.add(t.anchor)
        if glued is not None:
            out.append((t, t, glued))
            continue
        for c in grid.right_of(t, span=span):
            if c.value not in (None, ""):
                out.append((t, c, c.value))
                break
    return out


#: 表示形式の中の**引用された文字**（画面に文字として出る所）。
#:   例: `"請求日： "yyyy\年m\月d\日;@` → `請求日： `
_FMT_LITERAL = re.compile(r'"([^"]*)"')


def format_label(fmt) -> str:
    """表示形式に焼き込まれた文字だけを取り出す（書式コードは落とす）。

    ★★ なぜ要るか（2026-09-12 の実測）: 実物の雛形には、`請求日` `請求書番号` が
      **セルの文字ではなく表示形式**に入っているものが在る。画面には
      「請求日： 2026年6月30日」と見えるのに、セルの文字を読む器官には
      **ラベルが 1 文字も存在しない**（セルの中身は日付型・裸の数値だけ）。
      検体 87 冊のうち 13 冊・実物の雛形にも在る形で、珍しい癖ではない。
    """
    return "".join(_FMT_LITERAL.findall(str(fmt or "")))


def _labelled_by_format(grid: Grid, labels) -> list:
    """表示形式にラベルが焼き込まれたセル。戻り値: [(ラベルの語, セル)]

    ★ 値はそのセル自身（ラベルの右ではない ── ラベルはセルに貼り付いている）。
    ★ 拾いすぎの危険は測ってある: ラベル語を含む表示形式は、全 180 冊で
      日付型／数値のセルにしか付いていなかった（2026-09-12）。
      それでも値が読めなければ、下の読み手が rivals として理由を残す。
    """
    out, seen = [], set()
    for c in grid.all_cells():
        if c.anchor in seen or not c.fmt:
            continue
        lit = norm(format_label(c.fmt))
        if not lit:
            continue
        for w in labels:
            if norm(w) in lit:
                seen.add(c.anchor)
                out.append((w, c))
                break
    return out


def _looks_like_placeholder_date(raw) -> bool:
    return isinstance(raw, str) and any(ch in raw for ch in _DATE_PLACEHOLDER_CHARS)


def _labels_with_nothing_right(grid, labels, *, span: int = 4) -> list:
    """ラベルは在るのに**右が空**のセル（値が読めなかった理由を人に返すため）。

    ★★ なぜ別の器か（2026-09-13 の実測）: 合成検体 87 冊で請求日が空だった 74 冊のうち
      **61 冊はラベルが在って値が空**だった。なのに理由は「請求日が見つかりませんでした
      （…ラベルの右か同居だけ）」だけで、**ラベルの在処を言っていなかった**（50 冊）。
      買い手にとって「うちの請求書に日付が入っていない」と「この道具が読めなかった」は
      **別の話**で、後者だと思われると道具が疑われる。
    ★ `_labelled_text_right` は「右が空」を返さない（下に落ちない設計の副作用）。
      共有の器は**変えない** ── 宛先・請求元も通るので血流が広い。ここは
      **理由を作る枝だけ**で呼ぶ新しい小さな器にする。
    ★ 読みの規則は 1 文字も変えない（値は 1 つも増えない）。
    """
    wanted = {norm(x).rstrip("：:") for x in labels}
    out, seen = [], set()
    for t in grid.text_cells():
        if t.anchor in seen or norm(str(t.value)).rstrip("：:") not in wanted:
            continue
        if any(c.value is not None and str(c.value).strip() for c in grid.right_of(t, span)):
            continue
        seen.add(t.anchor)
        out.append(t)
    return out


def _empty_label_note(grid, labels, field: str) -> str:
    """「ラベルは在ったが右は空」を 1 文に。★ 無いときは空文字（余計な文を足さない）。"""
    cells = _labels_with_nothing_right(grid, labels)
    if not cells:
        return ""
    where = "／".join(f"{c.at}「{norm(c.value)[:8]}」" for c in cells[:3])
    return (f": {where} のラベルは在りましたが、その右は空でした"
            f" ── この帳票には{field}が入っていないようです")


def read_issue_date(grid: Grid) -> Record:
    """請求日。★ ラベルの右か同居だけ。読めない形は値を作らず、何を探したかを言う。"""
    found = [(lab, cell, raw) for lab, cell, raw in _labelled_text_right(grid, _LABEL_DATE)]
    # ★ ラベルが表示形式に焼き込まれている冊（実物 inv21 の癖）── セルの文字には無い
    found += [(None, cell, cell.value) for w, cell in _labelled_by_format(grid, _LABEL_DATE)
              if not any(c.anchor == cell.anchor for _l, c, _r in found)]
    evid, rivals = [], []
    for lab, cell, raw in found:
        if _looks_like_placeholder_date(raw):
            rivals.append((cell.at, str(raw)[:16], "雛形の埋め草のままで、日付ではありません"))
            continue
        # ★ 受ける暦を決めるのはここ（算術は date_compare が 1 つだけ持つ）。
        #   和暦は帳票に実在する形で、元号は法で決まる閉じた一覧 ── 言い回しの列挙ではない。
        d = _parse_date(raw) or _parse_wareki(raw)
        if d is None:
            rivals.append((cell.at, str(raw)[:16], "日付として読めません"))
            continue
        how = (f"{lab.at}「{norm(lab.value)[:8]}」の右 {cell.at}" if lab is not None
               else f"{cell.at}（表示形式に「{norm(format_label(cell.fmt))[:8]}」）")
        evid.append(Evidence(rule="請求日の欄", value=d, at=cell.at, how=how))
    if not evid:
        why = ("請求日が見つかりませんでした（『請求日』『発行日』のラベルの右か、"
               "同じセルに西暦か和暦の日付が入っている形だけを読みます）")
        if rivals:
            why += ": " + "／".join(f"{at} は{note}" for at, _v, note in rivals)
        else:
            why += _empty_label_note(grid, _LABEL_DATE, "請求日")
        return Record("請求日", (), rivals=tuple(rivals), blank_reason=why)
    return _record("請求日", evid, rivals)


def read_invoice_number(grid: Grid) -> Record:
    """請求番号。★ ラベルの右か同居だけ（下には落ちない）。在れば強い識別子・無くても止まらない。"""
    found = [(lab, cell, raw) for lab, cell, raw in _labelled_text_right(grid, _LABEL_NUMBER)]
    found += [(None, cell, cell.value) for w, cell in _labelled_by_format(grid, _LABEL_NUMBER)
              if not any(c.anchor == cell.anchor for _l, c, _r in found)]
    evid = []
    for lab, cell, raw in found:
        # ★ 請求番号は**識別子**であって数ではない（合計しない・先頭 0 を落とさない）ので
        #   文字として出す。裸の数値セル（実物 inv21）が 20260601.0 に化けないよう均す。
        if isinstance(raw, float) and raw.is_integer():
            raw = int(raw)
        txt = str(raw).strip()
        if not txt or _looks_like_placeholder_date(txt):
            continue
        # ★★ 請求番号は識別子 ── **数字を含まないものは識別子ではない**（2026-09-12）。
        #   素の `No` をラベルに足したので、明細の見出しの右（`品目`）を弾く構造の規則が要る。
        #   実測: いま出している請求番号に数字を含まないものは 0 件（束 51・検体 13・入れ子 2）。
        if not _HAS_DIGIT.search(txt):
            continue
        how = (f"{lab.at}「{norm(lab.value)[:8]}」の右 {cell.at}" if lab is not None
               else f"{cell.at}（表示形式に「{norm(format_label(cell.fmt))[:8]}」）")
        evid.append(Evidence(rule="請求番号の欄", value=txt, at=cell.at, how=how))
    if not evid:
        return Record("請求番号", (), blank_reason=(
            "請求番号が見つかりませんでした（『請求番号』『請求書No』のラベルの右か、"
            "同じセルに番号が入っている形だけを読みます）"
            + _empty_label_note(grid, _LABEL_NUMBER, "請求番号")))
    return _record("請求番号", evid, [])


#: ★ 「請求額」がどの数字を指すか決められなくなる欄。
#:   2026-09-11 の実測から ── 繰越請求（前回請求額・入金額・繰越金額・今回請求額）や
#:   源泉徴収の控除（合計金額とお振込金額が違う）が在る帳票では、
#:   合計を「請求額」として出すと**人が振り込む額と違う数字**を渡すことになる。
#:   ★ こういう冊は「取れなかった」のではなく「**決められない**」。空欄＋理由が正解。
#:   ★ 実物は `B42=お振込金額　121,790 円` のように **ラベルと金額が同じセルに同居**する
#:     ので、ラベルの右を見る規則では一生見つからない。文字として探す。
_AMBIGUITY_MARKS = ("繰越", "前回請求額", "源泉", "差引請求", "差引金額",
                    "お振込金額", "御振込金額", "相殺", "前受金")

#: 帯の始まりを示す語。★ この語がどこかに在る行から下は、もう明細ではない。
_BAND_LABELS = {norm(x) for x in (
    "小計", "小計金額", "税抜金額", "消費税", "消費税額", "内消費税", "税額",
    "合計", "合計金額", "総合計", "税込合計", "ご請求金額", "請求金額",
    "10%対象", "8%対象", "10％対象", "8％対象", "対象額（税抜）", "税率区分",
    # ★★ `値引`/`値引き` は 2026-09-12 に**外した**（器官の初版が a priori に入れていた語で、
    #   実測の事件に紐づいていなかった）。帯は 小計/合計/消費税 で開くので、**それより前に
    #   出る値引きは明細の一部**だ ── 走査は最初の帯の語で止まるため、値引が効くのは
    #   「最も早い帯の語」の時、つまり明細行として在る時だけだった。
    #   実測 T15（値引き行 −5,000 の陰性対照）: 明細から落ちて 110,000 になり、小計 105,000 と
    #   合わない**偽の食い違い**を出していた。どちらに数えても小計に足し合わされる側の行。
    "備考")}

#: 明細の見出し。★ この行が見つかれば、その下が明細ブロック。
_DETAIL_HEAD_ITEM = ("品番・品名", "品名", "品目", "内容", "摘要", "件名", "作業内容")
_DETAIL_HEAD_AMOUNT = ("金額", "金額（税抜）", "小計",
                       # ★ 2026-09-12: 実物の骨 spread_1/3/5 は行ごとの金額の列が **合計** と名乗る
                       #   （`商品コード | 品目 | 数量 | 単位 | 単価 | 合計`）── 検体 85 冊のうち 11 冊。
                       #   帯の「合計」と同じ語だが、見出し行は品名側の語を同じ行に持つのが条件なので
                       #   帯の行とは混ざらず、帯の走査は小計で止まる。実測: 見出し 72→84・score 不変・
                       #   単→確 10 冊（明細の合計が小計と一致）・over-claim 0・PDF 不変。
                       "合計")


#: 明細の計算に関わる列の見出し。★ ここが空だと、合計は正しく見えても中身が抜ける。
_COL_QTY = ("数量", "数 量", "数")
_COL_UNIT = ("単価", "単 価", "単価（税抜）")


def detail_row_anomalies(grid: Grid, item_head, amt_head, stop_row: int) -> list:
    """明細の**行ごとの算術**を見る。★ 帯と合計が一致していても中身は壊れている。

    ★★ 2026-09-11、未見 51 冊の実測から:
      数量の入力漏れで `単価 5,000 × 数量(空) = 金額 0` になっている行があった。
      小計は SUM なので **0 を足しても一致する** ── 上部・帯・明細合計の
      3 つが揃って一致し、「確」が立つ。**壊れた冊が満場一致で通る。**
      → 行の中の算術（数量 × 単価 ＝ 金額）を見に行く。

    戻り値: [(行, 事情の 1 行)]
    """
    # 見出しの行から、数量・単価の列を拾う
    qty_col = unit_col = None
    for c in grid.all_cells():
        if c.row != item_head.row or not isinstance(c.value, str):
            continue
        t = norm(c.value)
        if qty_col is None and t in {norm(x) for x in _COL_QTY}:
            qty_col = c.col
        if unit_col is None and t in {norm(x) for x in _COL_UNIT}:
            unit_col = c.col
    if qty_col is None or unit_col is None:
        return []

    out = []
    for r in _detail_rows(grid, item_head, amt_head, stop_row):
        q = grid.cell(r, qty_col)
        u = grid.cell(r, unit_col)
        a = grid.cell(r, amt_head.col)
        qv = q.value if q is not None and _is_number(q.value) else None
        uv = u.value if u is not None and _is_number(u.value) else None
        av = a.value if a is not None and _is_number(a.value) else None
        label = _detail_row_name(grid, item_head, amt_head, r)
        if uv is not None and qv is None:
            out.append((r, f"{r} 行目「{label}」は単価 {uv:,.0f} が入っているのに"
                           f"数量が空で、金額が {0 if av is None else av:,.0f} です"))
            continue
        if qv is not None and uv is not None and av is not None:
            if abs(qv * uv - av) >= 0.005:
                out.append((r, f"{r} 行目「{label}」は 数量 {qv:g} × 単価 {uv:,.0f} と"
                               f"金額 {av:,.0f} が合いません"))
    return out


def _detail_rows(grid: Grid, item_head, amt_head, stop_row: int) -> list:
    """明細の行を数える。★ 行の目印は**金額の列に数が在ること**。

    ★ 初版は「品名の列に文字が在る行」で数えていたが、実物は見出しと値の列が
      ずれる（inv21 は見出し `C32=品目` に対して値が `D33`）。
      その結果、行が 1 つも見つからず番人が黙って何も言わなかった。
    """
    return [r for r in range(amt_head.row + 1, stop_row)
            if grid.cell(r, amt_head.col) is not None
            and _is_number(grid.cell(r, amt_head.col).value)]


def _detail_row_name(grid: Grid, item_head, amt_head, row: int) -> str:
    """その行の品名らしい文字（見出しの列から金額の列の手前まで探す）。"""
    for c in range(item_head.col, amt_head.col):
        got = grid.cell(row, c)
        if got is not None and isinstance(got.value, str) and got.value.strip():
            return got.value.strip()[:14]
    return f"{row} 行目"


def detail_blank_columns(grid: Grid, item_head, amt_head, stop_row: int) -> list:
    """明細の中で「**他の行は埋めているのに、この行だけ空**」の欄を挙げる。

    ★★ 2026-09-11、未見 51 冊の実測から:
      小計が `SUMIF(J33:J48, 10%, K33:K48)` で、J34（税率）だけが空だったため
      その行の 15,000 が丸ごと落ちていた。合計は「正しく計算されて」いる ──
      **拾われなかっただけ**。
      ★ 式を辿って SUMIF の条件列を割り出すこともできるが、それは式の書き方に
        依存する。「他の行が埋めている欄がこの行だけ空」なら、式が何であれ怪しい。
        こちらの方が言い方として正直で、しかも壊れ方に依らない。

    戻り値: [(行, 見出しの語)]
    """
    headers = {}
    for c in grid.all_cells():
        if c.row == item_head.row and isinstance(c.value, str) and c.value.strip():
            headers[c.col] = norm(c.value)
    rows = _detail_rows(grid, item_head, amt_head, stop_row)
    if len(rows) < 2:
        return []

    out = []
    for col, name in headers.items():
        if col == item_head.col:
            continue
        filled = [r for r in rows if grid.cell(r, col) is not None
                  and str(grid.cell(r, col).value).strip() not in ("", "None")]
        # ★ 半分以上の行が埋めている欄だけを見る（備考のような任意欄で騒がない）
        if len(filled) * 2 <= len(rows) or len(filled) == len(rows):
            continue
        for r in rows:
            if r not in filled:
                out.append((r, name))
    return out


def detail_amount_sum(grid: Grid, band_row: int | None = None):
    """明細ブロックの「金額」列を足す。★ 帯とは**別の口**。

    ★★ なぜ要るか（2026-09-11、未見 51 冊の実測）:
      初版は上部の請求額欄と帯の合計しか見ず、両者が一致すると「確」と言った。
      だが実物の壊れ方は **明細と帯のあいだ**で起きる ──
        ・集計範囲の外に明細が追記されている（式は正しいのに合計が足りない）
        ・税率列の入力漏れで SUMIF がその行を落とす
        ・数量の入力漏れで金額が 0 になる行
      どれも上部と帯だけ見ていると**一致していて気づけない**。

    戻り値: (合計, 品名の見出し, 足したセル, 金額の見出し, 明細の終わりの行)
    """
    # ★★ 列を座標から組み直した格子では、この表歩きをしない（2026-09-12 の実測）。
    #   推定した列の上を歩くと隣の列（単価・数量）を足しうる ── PDF で出た食い違い 39 件の
    #   うち **34 件が偽**だった（Excel では裏が取れて正しい値が出ている冊）。
    #   ★ 「空欄は誤値より安い」は「偽の疑いも安い」を意味しない ── オオカミ少年は
    #     この製品がいちばん避けるもの。掃けないなら掃けないと言って降りる。
    if grid.columns_reconstructed:
        return None, None, (), None, 0

    head_item = {norm(x) for x in _DETAIL_HEAD_ITEM}
    head_amt = {norm(x) for x in _DETAIL_HEAD_AMOUNT}
    best = None
    for t in grid.text_cells():
        if norm(t.value) not in head_item:
            continue
        # 同じ行に「金額」の見出しが在るか
        for c in grid.right_of(t, span=grid.cols):
            if norm(c.value) in head_amt:
                best = (t, c)
                break
        if best:
            break
    if not best:
        return None, None, (), None, 0

    item_head, amt_head = best
    stop = band_row if band_row else grid.rows
    # ★★ 明細は「小計ラベルの行まで」ではなく「**帯が始まる行まで**」（2026-09-11 の実測）。
    #   construction_bill は小計 C28 の手前に税区分の行（10%対象 C26 / 8%対象 C27）が在り、
    #   そこまで明細として足すと **区分ごとの消費税額まで明細に混ざる**。
    #   実測 B24: 明細 117,000 に G26=11,700 が乗って 128,700 ＝ 税込合計と一致してしまい、
    #   「明細と小計が合わない」という**偽の食い違い**が出た（数が合うので気づきにくい）。
    for r in range(amt_head.row + 1, stop):
        if any(norm(c.value) in _BAND_LABELS for c in grid.all_cells()
               if c.row == r and isinstance(c.value, str)):
            stop = r
            break

    cells, blanks = [], 0
    for r in range(amt_head.row + 1, stop):
        got = grid.cell(r, amt_head.col)
        name = grid.cell(r, item_head.col)
        has_name = name is not None and isinstance(name.value, str) and name.value.strip()
        if got is not None and _is_number(got.value):
            cells.append(got)
            blanks = 0
        elif has_name:
            blanks = 0                       # ★ 品名は在るが金額が空 ── 明細は続いている
        else:
            blanks += 1
            if blanks >= 4:                  # ★ 空行 1 つで切らない（実物は 1 行空く）
                break
    if not cells:
        return None, item_head, (), amt_head, stop
    # ★ 同じアンカーを 2 度足さない（結合の写しで二重計上する）
    seen, total = set(), 0.0
    for c in cells:
        if c.anchor in seen:
            continue
        seen.add(c.anchor)
        total += float(c.value)
    return total, item_head, tuple(cells), amt_head, stop


# ── 依頼された 1 項目ずつの規則 ─────────────────────────────
def read_addressee(grid: Grid) -> Record:
    """宛先（請求を受ける側）。

    ★ 癖①への処方 ── 御中/様 の付いたセルを見つけたら:
        そのセル自身が組織の形をしていれば、敬称を落としたものが名前。
        そうでなければ（部署名・担当者名だった）、**同じ列の直上で最初に出会う文字**。
    """
    evid, rivals = [], []
    for t in grid.text_cells():
        raw = str(t.value)
        if not any(h in raw for h in _HONORIFIC):
            continue
        # ★ 名前を作るのは `clean_org_name` 一本（プレースホルダの番人もその中）
        name, why = clean_org_name(raw)
        if name:
            evid.append(Evidence(rule="敬称と同じセル", value=name, at=t.at,
                                 how=f"{t.at} の「{raw[:18]}」から敬称を除いた"))
            continue
        # 敬称のセルは部署か担当者 ── 名前はその上
        above = _nearest_text_above(grid, t)
        if above is None:
            rivals.append((t.at, raw[:24], f"敬称は在るが、{why}／上にも文字がない"))
            continue
        name2, why2 = clean_org_name(above.value)
        if name2:
            evid.append(Evidence(rule="敬称の直上", value=name2, at=above.at,
                                 how=f"{t.at} の「{norm(raw)[:12]}」の上、{above.at}"))
        else:
            rivals.append((above.at, str(above.value)[:24], why2))

    if not evid:
        why = "「御中」「様」の付いた宛名が見つかりませんでした"
        if rivals:
            why = ("宛先を決められませんでした: "
                   + "／".join(f"{at} は{note}" for at, _v, note in rivals[:3]))
        return Record("宛先", (), rivals=tuple(rivals), blank_reason=why)
    return _record("宛先", evid, rivals)


def _nearest_text_above(grid: Grid, cell, span: int = 6):
    """同じ列を上へ辿って最初に出会う文字セル（空行は跨ぐ）。"""
    for r in range(cell.row - 1, max(0, cell.row - span - 1), -1):
        got = grid.cell(r, cell.col)
        if got is None or got.anchor == cell.anchor:
            continue
        if isinstance(got.value, str) and got.value.strip():
            return got
    return None


def read_issuer(grid: Grid, addressee: Record) -> Record:
    """請求元（請求を出す側）。

    ★ 癖②③への処方 ── 「〒の上」のような**骨に依存する位置**では決めない。
      組織の形をした名前をすべて拾い、次で絞る:
        ① 宛先の名前を**含む**候補は除く（同じ形の 2 ブロックの片方を消す）
        ② 登録番号（T+13 桁）が同じ列の近くに在るものを優先する
           ★ 登録番号は請求を**出す**側にしか付かない。これは「どちらのブロックが
             発行者か」という**役割**の証拠であって、名前そのものの裏取りではない
        ③ 残った候補は**全部そのまま根拠にする** ── 1 つなら 単、値が違う 2 つなら 割。
           区分の導出（grade_of）が既に知っている仕事なので、ここでは決めない。

    ★★ 2026-09-11（B′）: ③ は以前「右側／上側を優先する」だった。
      これは実物の版面の**癖**であって根拠ではない ── 当たらない冊では
      **黙って間違った名前**を出す。実測で何冊がその癖に頼っていたかを測った:

          候補が 1 つだけ      78 冊   ← 癖は一度も走らない
          癖で決めていた        2 冊   ← T10（宛先の再掲）・T11（宛先が 割 で除けない）
          候補なし              3 冊
          シートが決まらない    4 冊

      ① を「完全一致 → 含む」に変えると T10 の競合が構造的に消え、**79 冊が候補 1 つ**に
      なる（正解を巻き添えで消した冊は 0）。残る T11 は宛先そのものが 割 なので
      買い手を除けない ── そこは決められないのが正直で、検体も期待を宣言していない。
      → 癖を消しても失うものが無いことを測ってから消した。

    ★ 正直に残す穴（この検体では測れない）: ① の「含む」は、**社名が入れ子**のときに
      正解を巻き添えにしうる（宛先「トヨタ自動車」／請求元「トヨタ自動車東日本」）。
      いまの検体に入れ子の社名は 1 つも無いので、測れていない。
      ★ 発火条件: 入れ子の社名を持つ検体が入った時、または「請求元が空欄になる」報告が来た時。
    """
    taken = norm(grade_value(addressee) or "")

    cands, dropped, seen = [], [], set()
    same_as_addressee: list = []          #: ★ 宛先と同じで捨てた候補（理由に名指しする）
    for t in grid.text_cells():
        if t.anchor in seen:
            continue
        raw = str(t.value).strip()
        if not _looks_like_org(raw):
            continue
        if any(h in raw for h in _HONORIFIC):
            continue                                   # 敬称つきは受け手の側
        seen.add(t.anchor)
        # ★ 名前を作るのは `clean_org_name` 一本 ── 宛先の側と同じ関数を通す。
        #   （片方にだけ番人を書くと、実物の空雛形で片側だけが値を出す）
        name, why = clean_org_name(raw)
        if not name:
            dropped.append((t.at, raw[:26], why))
            continue
        # ① 宛先を**含む**候補は採らない（完全一致では飾り付きの再掲がすり抜ける）。
        #   実測 T10: 『御請求先：ナギ商会株式会社／9 月分』が候補に残っていた。
        #   ★★ 2026-09-12: ここで黙って `continue` していたせいで、**理由が嘘をついていた** ──
        #     実物の請求書（買い手と売り手に同じ社名が入っている形）で、法人格つきの名前を
        #     見つけて捨てたのに「『株式会社』などの法人格が付いた名前だけを探しています ──
        #     屋号や略称だけの請求書では見つかりません」と言っていた。
        #     人は探し方を疑うが、本当に疑うべきは「宛先と同じ名前と判定された」の方だ。
        #   → 捨てた候補は**記録して名指しする**（捨てる判断は変えない）。
        if taken and taken in norm(name):
            same_as_addressee.append((t.at, name, "宛先と同じ名前です（請求元と宛先が"
                                      "同じ名前に見えるので、どちらが発行元か決められません）"))
            continue
        cands.append((t, name))

    if not cands:
        # ★ 何を探したのかを言う（2026-09-11・入れ子の検体 14 冊の実測から）。
        #   「見つかりませんでした」だけだと、人は**探し方**を疑えない。実際に落ちたのは
        #   `ナギ商会`（略称）・`ナギ商店`（屋号）・ブランド名 ── どれも法人格が無いので
        #   候補にすら上がらなかった。個人事業主と屋号は実務では普通にある。
        #   ★ 拾いに行く方は測って見送った（§3.7）: 位置や隣接では見出し・住所と
        #     見分けられず、除外語を足し続けない限り成り立たない。だから
        #     **できないことを正直に言う**方に倒す。
        # ★★ 捨てた候補は**全部**並べる（2026-09-12）── 分岐で片方を押しのけない。
        #   初版はここを「宛先と同じ」と「雛形のまま」で分岐させ、前者が後者を隠して
        #   T05（発行者名が `株式会社 〇〇〇`）の名指しを消した ── 同じ形の理由を
        #   2 通りに書くと、片方が片方を食う。1 本に畳む。
        discarded = tuple(same_as_addressee) + tuple(dropped)
        base = ("発行元（請求元）の名前が見つかりませんでした"
                "（『株式会社』などの法人格が付いた名前だけを探しています ── "
                "屋号や略称だけの請求書では見つかりません）")
        if not discarded:
            return Record("請求元", (), blank_reason=base)
        why = ("発行元（請求元）を決められませんでした ── 名前の候補は "
               f"{len(discarded)} 件見つかりましたが、どれも使えませんでした: "
               + "／".join(f"{at}「{str(v)[:14]}」は{note}" for at, v, note in discarded[:3]))
        return Record("請求元", (), rivals=discarded, blank_reason=why)

    if not taken:
        # ★★ 2026-09-12: 宛先が決まらなかった冊では、**買い手を候補から除けない**。
        #   請求書は請求元と宛先が「同じ形の 2 ブロック」で並ぶ（癖②③）ので、片方を
        #   同定できないまま名前を出すと、**買い手の名前を請求元として出しうる**。
        #   ★ 実測（PDF 化した 87 冊・2026-09-12）: この経路で 14 件、
        #     買い手『ナギ商会株式会社』を請求元として出していた ── 支払いの文脈で最悪の間違い。
        #   ★ この関数の docstring は以前から「宛先が 割 なので買い手を除けない ──
        #     そこは決められないのが正直」と**宣言していたのに、実体は値を出していた**
        #     （依頼/宣言/実体の三項のうち、宣言と実体の食い違い）。
        #   ★ Excel でも同じ経路は在った。宛先が一度も失敗しなかったから発火しなかっただけ
        #     （出ないことは信号でない）。処置の値段も測った: 影響は検体 152 冊中 1 冊（入09）で、
        #     その 1 冊は**そもそも拒否が正解**の冊だった。
        return Record("請求元", (), rivals=tuple(
            (t.at, n, "宛先が決まらないので、買い手の名前と見分けられない") for t, n in cands),
            blank_reason=(
                "宛先が決まらなかったので、請求元を決められませんでした"
                f"（組織の名前は {len(cands)} 件見つかっています: "
                + "／".join(f"{t.at} の「{n[:14]}」" for t, n in cands[:3])
                + "。請求書は請求元と宛先が同じ形で並ぶので、宛先を同定できないまま"
                  "名前を出すと、買い手の名前を請求元として出すおそれがあります）"))

    with_reg = [(t, n) for t, n in cands if _has_regno_near(grid, t)]
    chosen = with_reg or cands
    how = ("登録番号が近くに在る" if with_reg else "組織名で、宛先ではない")
    # ③ 残った候補を**全部**根拠にする ── 選ばない。1 つなら 単、値が違えば 割。
    evid = [Evidence(rule="請求元の名前", value=n, at=t.at,
                     how=f"{t.at} の「{n[:18]}」（{how}）") for t, n in chosen]
    rivals = [(t.at, n, "登録番号が近くに無いので採らなかった候補")
              for t, n in cands if (t, n) not in chosen]
    return _record("請求元", evid, rivals)


def _has_regno_near(grid: Grid, cell, span: int = 3) -> bool:
    """登録番号（T+13 桁）が同じ列の上下 span 行に在るか。"""
    for r in range(max(1, cell.row - span), min(grid.rows, cell.row + span) + 1):
        for c in (cell.col, cell.col + 1):
            got = grid.cell(r, c)
            if got is None:
                continue
            if _REGNO.match(norm(got.value)):
                return True
            if "登録番号" in norm(got.value):
                return True
    return False


def read_billed_total(grid: Grid, ws_formula=None) -> Record:
    """請求額（税込の総額）。★ 独立した口を集めて `grade()` に渡す。

    口:
      ① 上部の「ご請求金額」欄
      ② 帯の「合計金額」
      ③ 小計 ＋ 消費税  ← ★ 合計が式でそれを足しているなら**独立ではない**
    """
    # ★★ 帯の在り処を「シートの下半分」で決めない（2026-09-11、未見 51 冊で踏んだ）。
    #   construction_bill は帯が 26〜28 行目に在り、60 行の格子の下半分（30 行目以降）
    #   から漏れて、**帯そのものを見落としていた**。
    #   → 位置ではなく**構造**で切る: 明細の見出しより下が帯、より上が表書き。
    _d_sum0, d_head0, _c0, _a0, _s0 = detail_amount_sum(grid, None)
    split_row = d_head0.row if d_head0 is not None else max(1, grid.rows // 2)
    half = split_row

    evid, rivals, excluded = [], [], []
    #: ★ 読めなかった枠の事情。空欄にするときの理由になる。
    unreadable: list = []

    for lab, val, note in _labelled_number(grid, _LABEL_BILLED, rows=(1, half)):
        if note:
            unreadable.append(note)
            break
        evid.append(Evidence(rule="上部の請求額欄", value=val.value, at=val.at,
                             how=f"{lab.at}「{norm(lab.value)[:12]}」の右 {val.at}"))
        break

    totals = _labelled_number(grid, _LABEL_TOTAL, rows=(half, grid.rows))
    total_cell = None
    for lab, val, note in totals:
        if note:
            unreadable.append(note)
            break
        total_cell = val
        evid.append(Evidence(rule="帯の合計", value=val.value, at=val.at,
                             how=f"{lab.at}「{norm(lab.value)[:12]}」の右 {val.at}"))
        break

    sub = _labelled_number(grid, _LABEL_SUBTOTAL, rows=(half, grid.rows))
    tax = _labelled_number(grid, _LABEL_TAX, rows=(half, grid.rows))
    if sub and tax and not sub[0][2] and not tax[0][2]:
        s_at, t_at = sub[0][1], tax[0][1]
        # ★★ 帯には 2 つの形が在る（2026-09-11、税率群で踏んだ）:
        #   積み上げ型: `小計`→H37 / `消費税`→H38 ── 値が**同じ列**に縦に並ぶ
        #   表型      : `対象額（税抜）| 消費税` が列見出しで、
        #               `10%対象 / 8%対象 / 小計` が行 ── 小計の**行**の消費税列を見る
        #   初版は表型で 10% 区分の税額を小計に足し、偽の食い違いを出した。
        #   ★ 見分け: 小計の値と消費税の値が違う列なら表型。
        if t_at.col != s_at.col:
            same_row = grid.cell(s_at.row, t_at.col)
            if same_row is not None and _is_number(same_row.value):
                t_at = same_row
        got = float(s_at.value) + float(t_at.value)
        if total_cell is not None and _depends_on(ws_formula, total_cell, (s_at, t_at)):
            # ★ 合計が小計＋消費税の式そのもの ── 一致して当たり前。数えない。
            excluded.append((f"{s_at.at}+{t_at.at}", got,
                             f"{total_cell.at} の式が {s_at.at} と {t_at.at} を"
                             "足しているので、検算になりません"))
        else:
            evid.append(Evidence(rule="小計＋消費税", value=got,
                                 at=f"{s_at.at}+{t_at.at}",
                                 how=f"小計 {s_at.at} と消費税 {t_at.at} の和"))

    # ── ★ 掃き出し①: そもそも「請求額」が一意に決まる帳票か ──────
    # ★ 見つけた欄を**全部**並べる。最初の 1 つだけ引くと、
    #   「前回請求額」は名指せても「繰越金額」を名指せず、人は何が起きたか分からない。
    found, seen_mark = [], set()
    for t in grid.text_cells():
        raw = str(t.value)
        for m in _AMBIGUITY_MARKS:
            if m in raw and m not in seen_mark:
                seen_mark.add(m)
                found.append(f"「{norm(raw)[:20]}」（{t.at}）")
    if found:
        why = (f"この請求書には {'・'.join(found)} の欄があり、"
               "合計と実際にお支払いいただく額が違います。"
               "どちらを請求額とすべきか決められないので、空欄にしました")
        return Record("請求額", tuple(evid), tuple(rivals), tuple(excluded),
                      why, False, "", True, why)

    # ── ★ 掃き出し②: 明細ブロックと帯が合っているか ──────────────
    #   一致は「他の口が黙っている」ことを意味しない。ここで**開きに行く**。
    swept, swept_how = False, ""
    band_row = sub[0][0].row if sub else (total_cell.row if total_cell else None)
    d_sum, d_head, _d_cells, d_amt, d_stop = detail_amount_sum(grid, band_row)
    if d_sum is not None and sub and not sub[0][2]:
        s_cell = sub[0][1]
        if abs(d_sum - float(s_cell.value)) >= 0.005:
            # ★ 式が正しくても範囲が足りないことがある（集計範囲外の追記・SUMIF の漏れ）。
            #   式かどうかに関係なく、食い違いは食い違い。
            gaps = detail_blank_columns(grid, d_head, d_amt, d_stop)
            # ★ 「合わない」だけでは人は直せない。**どの行の何が空か**まで言う。
            tail = ("" if not gaps else
                    "（" + "・".join(f"{r} 行目の「{n}」が空です" for r, n in gaps[:3])
                    + "。ここが合計から漏れている可能性があります）")
            why = (f"明細の金額の合計（{d_sum:,.0f}）と小計 {s_cell.at}"
                   f"（{float(s_cell.value):,.0f}）が合いません{tail}。"
                   "どちらが正しいか決められないので、請求額は空欄にしました")
            # ★ 食い違いは `conflict` として**導出に渡す**。理由の文字列にだけ書くと、
            #   区分は evidences しか見ないので「単・値あり・理由は食い違い」が出る。
            return Record("請求額", tuple(evid), tuple(rivals), tuple(excluded),
                          why, False, "", True, why)
        # ★ 合計が合っていても中身は壊れていることがある（0 を足しても SUM は合う）。
        bad = detail_row_anomalies(grid, d_head, d_amt, d_stop)
        if bad:
            why = ("明細に、金額が正しく計算されていない行があります: "
                   + "／".join(t for _r, t in bad)
                   + "。合計を信じてよいか決められないので、請求額は空欄にしました")
            return Record("請求額", tuple(evid), tuple(rivals), tuple(excluded),
                          why, False, "", True, why)
        swept = True
        swept_how = f"明細 {d_head.at} 以下の合計とも一致（{d_sum:,.0f}）"
    elif d_sum is None:
        swept_how = "明細ブロックが見つからず、突き合わせできていません"

    # ★ 枠は埋まっているのに読めなかった ── これは「見つからない」ではない。
    #   黙って別のセルの数を返すより、読めないと言う方が安い。
    if unreadable:
        return Record("請求額", (), rivals=tuple(rivals), excluded=tuple(excluded),
                      blank_reason="請求額を読み取れませんでした: " + "／".join(unreadable))
    if not evid:
        return Record("請求額", (), rivals=tuple(rivals), excluded=tuple(excluded),
                      blank_reason="請求額の欄も、明細の帯の合計も見つかりませんでした")
    # ★ 裏取り済みを名乗れない理由を集める（1 箇所）── 写しを見分けられない（D4）と、
    #   掃き出していない口が在る（G3）。どちらも「一致した」までで止める。
    blind, blind_why = copies_are_indistinguishable(ws_formula)
    tentative = _record("請求額", evid, rivals, excluded, swept=swept, swept_how=swept_how)
    unconfirmable = (((blind_why,) if blind else ())
                     + (("明細の表は列を座標から組み直しているので、明細の合計は"
                         "当てにできません（掃き出していません）",)
                        if grid.columns_reconstructed else ())
                     + unswept_mouths(grid, grade_value(tentative)))
    return _record("請求額", evid, rivals, excluded, swept=swept, swept_how=swept_how,
                   unconfirmable=unconfirmable)


_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d{1,5})")


def _depends_on(ws_formula, cell, sources) -> bool:
    """`cell` の**式**が `sources` のセルを参照しているか（★ 恒真の検出）。"""
    if ws_formula is None:
        return False
    try:
        f = ws_formula.cell(row=cell.anchor[0], column=cell.anchor[1]).value
    except Exception:                                  # noqa: BLE001
        return False
    if not isinstance(f, str) or not f.startswith("="):
        return False
    refs = {f"{a}{b}" for a, b in _REF.findall(f.upper())}
    return any(s.at.upper() in refs for s in sources)


def copies_are_indistinguishable(ws_formula) -> tuple:
    """写し合いを見分ける手段が在るか。戻り値: (見分けられない, その 1 行)

    ★★ ここが唯一の判定（2026-09-12）。「PDF だから」ではなく
      「**式を読めないから**」で書く ── 次の入口（CSV・OCR）でも同じ 1 箇所で決まる。

    ★ なぜ要るか: 「別々の出所が一致した」を根拠に `裏が取れた` と言うには、
      その 2 つが写し合いでないと言えなければならない。Excel では式（`=合計`）を見て
      「この欄は帯の写しだ」と分かる（`_depends_on`）。式が読めない入口では言えない。
    """
    if ws_formula is not None:
        return False, ""
    return True, ("この帳票からは式を読めないので、同じ数字が別々に書かれたものか"
                  "写しかを見分けられません（PDF や、式を読み込まずに開いた表）")


#: 円でない通貨の印。★ `元` は入れない ── **『請求元』に当たる**（2026-09-12 の実測で誤爆）。
_FOREIGN_CURRENCY = ("USD", "EUR", "GBP", "CNY", "ドル", "ユーロ", "ポンド", "＄", "$", "€")

#: 帳票名そのものを名乗る**短いセル**。★ 語の有無では測れない ── misoca のフッタに
#:   「無料のクラウド見積・納品・請求書サービス」が在り、それで数えると偽陽性 29 冊（実測）。
_KIND_TITLE = re.compile(r"^(見積|納品|領収|受領)(書|明細)?"
                          r"(ESTIMATE|QUOTATION|DELIVERY|RECEIPT|INVOICE)?$", re.I)

#: 文章の中の金額（`合計金額 121,000 円（税込）をご請求申し上げます`）。
_NUM_IN_TEXT = re.compile(r"[0-9０-９][0-9０-９,，]{2,}")


def _numbers_in_text(text: str) -> list:
    out = []
    for m in _NUM_IN_TEXT.finditer(text):
        body = m.group(0).translate(str.maketrans("０１２３４５６７８９，", "0123456789,"))
        try:
            out.append(float(body.replace(",", "")))
        except ValueError:
            pass
    return out


def unswept_mouths(grid: Grid, total) -> tuple:
    """**掃き出していない口**の並び（1 行ずつ）。空でなければ裏取り済みを名乗らない。

    ★★ なぜ要るか（2026-09-12 の実測）: 明細の合計と一致しても、それは
      **明細という 1 つの口**を掃いただけだ。器官の初版が既に名指ししていた
      「開いていない第三の口」── 備考の金額の再掲・通貨が円でない・そもそも納品書か
      見積書か ── は別の口で、どれも「一致しているのに本体と違う」を作る。

    ★ 語の有無ではなく**構造**で測る（実測で偽陽性 0）:
        通貨    円でない通貨の印が在る（`元` は『請求元』に当たるので見ない）
        帳票名  **短いセル**が 見積書/納品書/領収書 を名乗る（長い文中の語は見ない）
        再掲    合計のラベルと金額が同じ文章に同居し、その金額が本体と違う

    ★ これらの冊が長く正しく見えていたのは、一文字の取りこぼし（`品番•品名` の U+2022）が
      掃き出し自体を壊していたから ── 番人が間違った理由で効いていた。
    """
    found = []
    for c in grid.text_cells():
        t = norm(c.value)
        hit = [k for k in _FOREIGN_CURRENCY if k in t]
        if hit:
            found.append(f"{c.at} に円でない通貨の印（{hit[0]}）があります")
            break
    for c in grid.text_cells():
        t = norm(c.value)
        if len(t) <= 12 and _KIND_TITLE.match(t):
            found.append(f"{c.at} が「{t}」と名乗っています（請求書ではない帳票が混ざっています）")
            break
    if isinstance(total, (int, float)):
        for c in grid.text_cells():
            t = norm(c.value)
            if len(t) <= 8 or not any(lbl in t for lbl in _LABEL_TOTAL):
                continue
            other = [x for x in _numbers_in_text(t) if abs(x - float(total)) > 0.5]
            if other:
                found.append(f"{c.at} に合計の再掲（{other[0]:,.0f}）があり、本体と違います")
                break
    return tuple(found)


def _record(field: str, evid: list, rivals: list, excluded: list = (),
            *, swept: bool = False, swept_how: str = "",
            unconfirmable: tuple = ()) -> Record:
    """`Record` を組む。★ 空欄になるなら**理由を必ず添える**（型が空を許さない）。

    ★ 区分は `grade_of` に聞く ── ここで `Record` を偽造して先読みすると、
      「区分は 1 箇所からしか作れない」という契約が骨抜きになる。
    """
    g = grade_of(evid, swept, unconfirmable=unconfirmable)
    if g in GRADES_WITH_VALUE:
        return Record(field, tuple(evid), tuple(rivals), tuple(excluded), "",
                      swept, swept_how, unconfirmable=unconfirmable)

    if g == SPLIT:
        # ★ 番地だけ並べても人には読めない ── **何の数字か**を書く
        #   （「E12=109999／L36=109999」では、どこを直せばいいか分からない）。
        seen, parts = set(), []
        for e in evid:
            if e.at in seen:
                continue
            seen.add(e.at)
            parts.append(f"{e.rule} {e.at}＝{e.value}")
        why = (f"根拠が食い違いました（{'／'.join(parts)}）。どれが正しいか"
               f"決められないので、{field}は空欄にしました")
    else:
        why = f"{field}の手がかりが見つかりませんでした"
    return Record(field, tuple(evid), tuple(rivals), tuple(excluded), why)


# ── 1 冊を読む ────────────────────────────────────────────
#: ★ この器官が返す項目。**ここが唯一の一覧**（forms_collect も read_book もこれを読む）。
#:   2026-09-11: read_book がシート未決/未発見の枝で自前の 3 つ組を持っていて、請求日・請求番号を
#:   足したとき**記録そのものが無い空欄**が生まれた（一覧は空欄・検分に理由なし）。
#:   「空欄には必ず理由」を型で守っていても、記録が無ければ型は働かない ── 一覧は 1 箇所に。
FIELDS = ("宛先", "請求元", "請求額", "請求日", "請求番号")


def read_form_grid(grid: Grid, ws_formula=None) -> dict:
    """1 つの格子を読んで、項目 → `Record` を返す。★ 区分は `field_record` が導く。

    ★ Excel のシートも PDF のページも、格子になればここは同じ（設計 D1）。
      `ws_formula` が無い入口（PDF・式を読まずに開いた表）では、請求額の裏取りを名乗らない（D4）。
    """
    addressee = read_addressee(grid)
    return {
        "宛先": addressee,
        "請求元": read_issuer(grid, addressee),
        "請求額": read_billed_total(grid, ws_formula),
        "請求日": read_issue_date(grid),
        "請求番号": read_invoice_number(grid),
    }


def read_form(ws, ws_formula=None, rows: int = 60, cols: int = 24) -> dict:
    """1 シートを読んで、項目 → `Record` を返す。"""
    return read_form_grid(Grid.read(ws, rows=rows, cols=cols, ws_formula=ws_formula), ws_formula)


#: 請求書らしさの印。★ どれか 1 つでは足りない（見積書にも合計は在る）。
def invoice_signals(grid: Grid) -> tuple:
    """そのシートが請求書らしいと言える**印**を並べる。

    ★ 1 つでは足りない ── 合計だけなら統計表にも在るし、
      社名だけなら送付状にも在る。**別の種類の印が 2 つ以上**で「らしい」とする。
    """
    texts = [norm(c.value) for c in grid.text_cells()]
    found = []
    if any(any(h in t for h in _HONORIFIC) for t in texts):
        found.append("宛名（御中・様）")
    if any("請求" in t for t in texts):
        found.append("「請求」の語")
    band = {lab for lab in ("小計", "消費税", "合計", "合計金額")
            if any(t == norm(lab) for t in texts)}
    if len(band) >= 2:
        found.append("帯（" + "・".join(sorted(band)) + "）")
    _s, head, _c, _a, _st = detail_amount_sum(grid, None)
    if head is not None:
        found.append(f"明細の見出し（{head.at}）")
    return tuple(found)


def read_book(wb, wb_formula=None, rows: int = 60, cols: int = 24) -> dict:
    """**1 冊**を読む。★ どのシートが請求書かも、ここで自分で決める。

    ★★ なぜ要るか（2026-09-11）:
      それまでの測定は、どのシートを読むかを**検体の答えから受け取っていた**。
      つまり「どのシートが請求書か」を一度も解いていないのに満点を出していた
      ── 買い手が持っていない手がかりで測っていた。
      実物の 14/87 冊は複数シートで、しかも:
        ・`説明` シートが 1 枚目で本体が 2 枚目（12 冊）
        ・1 冊に請求書が 2 枚（8 月分・9 月分）
        ・本体のシートが**非表示**で、可視シートは空の 1 枚だけ

    決め方:
      ① 各シートの「請求書らしい印」を数える（別種の印が 2 つ以上で候補）
      ② 可視の候補が 2 枚以上 → **どちらの請求書か決められない**ので全項目を空欄
      ③ 可視の候補が 1 枚 → それを読む
      ④ 可視に候補が無く、非表示に在る → 読むが、**非表示だったことを言う**
      ⑤ どこにも無い → 全項目を空欄（理由に、見たシート名を並べる）
    """
    seen, hidden, looked = [], [], []
    for name in wb.sheetnames:
        ws = wb[name]
        grid = Grid.read(ws, rows=rows, cols=cols)
        sig = invoice_signals(grid)
        looked.append(f"「{name}」({len(sig)} 印)")
        if len(sig) >= 2:
            (hidden if getattr(ws, "sheet_state", "visible") != "visible"
             else seen).append((name, sig))

    if len(seen) >= 2:
        names = "・".join(f"「{n}」" for n, _s in seen)
        why = (f"この 1 冊に請求書らしいシートが {len(seen)} 枚あります（{names}）。"
               "どちらの請求書か決められないので、空欄にしました")
        return {f: Record(f, (), (), (), why, False, "", True, why)
                for f in FIELDS}

    note = ""
    if seen:
        name = seen[0][0]
    elif hidden:
        name, _sig = hidden[0]
        note = f"（シート「{name}」は非表示でした）"
    else:
        why = ("請求書らしいシートが見つかりませんでした（見たシート: "
               + "・".join(looked) + "）")
        return {f: Record(f, (), (), (), why) for f in FIELDS}

    wsf = None
    if wb_formula is not None and name in wb_formula.sheetnames:
        wsf = wb_formula[name]
    recs = read_form(wb[name], ws_formula=wsf, rows=rows, cols=cols)
    if note:
        recs = {f: _with_note(r, note) for f, r in recs.items()}
    return recs


def read_pdf_book(path) -> dict:
    """PDF の **1 冊**を読む。★ どのページが請求書かも、ここで自分で決める（設計 D6）。

    決め方は `read_book`（Excel のシート選び）と同じ線:
      ① 各ページの「請求書らしい印」を数える（別種の印が 2 つ以上で候補）
      ② 候補が 2 ページ以上 → **どちらの請求書か決められない**ので全項目を空欄
      ③ 候補が 1 ページ → それを読む
      ④ どこにも無い → 全項目を空欄（理由に、見たページを並べる）
    ★ 答えから受け取る口は作らない（Excel 側で一度やって消した過ち）。
    ★ 明細が複数ページにまたがる形は**設計していない**（検体に 1 冊も無い・発火条件つき保留）。
    ★ テキスト層が無ければ `pdf_grid.NoTextLayer` がそのまま上がる ── 「請求書ではなかった」と
      混ぜない（設計 D7）。呼び出し側は読めなかった冊として名指しする。
    """
    from ailine_core import pdf_grid              # ★ pdfplumber の import は pdf_grid に閉じる

    seen, looked = [], []
    for page_no, grid in pdf_grid.grids_of(path):
        sig = invoice_signals(grid)
        looked.append(f"{page_no}頁({len(sig)} 印)")
        if len(sig) >= 2:
            seen.append((page_no, grid))
    if len(seen) >= 2:
        pages = "・".join(f"{n}頁" for n, _g in seen)
        why = (f"この 1 冊に請求書らしいページが {len(seen)} 枚あります（{pages}）。"
               "どちらの請求書か決められないので、空欄にしました")
        return {f: Record(f, (), (), (), why, False, "", True, why) for f in FIELDS}
    if not seen:
        why = ("請求書らしいページが見つかりませんでした（見たページ: "
               + "・".join(looked) + "）")
        return {f: Record(f, (), (), (), why) for f in FIELDS}
    return read_form_grid(seen[0][1], None)


def _with_note(rec: Record, note: str) -> Record:
    """人に見せる 1 行に一言足す（区分は変えない）。"""
    return Record(rec.field, rec.evidences, rec.rivals, rec.excluded,
                  (rec.blank_reason + note) if rec.blank_reason else "",
                  rec.swept, (rec.swept_how + note) if rec.swept_how else note,
                  rec.conflict, rec.conflict_why)
