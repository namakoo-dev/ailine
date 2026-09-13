"""forms_suspect — 帳票の**束**を見て初めて分かる怪しさを、所見として出す（2026-09-11）。

★★ 需要①の売り物は「代わりに疑うこと」。1 冊ずつは正常でも、同じ取引先が毎月出る束では
  「同じ請求書が 2 通」「同じ月に金額の違う 2 通」「1 冊だけ年が違う」「桁が違う」が見える。

★ ここは**値を作らない**。一覧の値は `field_record.value()` が決めたものをそのまま受け取り、
  所見（種類・関わる冊・理由）だけを返す。疑いは判断ではなく、人が見る場所の指さし。

★ 入力は「冊の名 → 項目 → 値」（割/無 は None）。値が None の項目はその冊の証拠にしない
  ── 読めなかったものを根拠に疑わない（空欄は誤値より安い、はここでも同じ）。

★ 取引先（請求元）が読めない冊は、取引先ごとの照合から外れる。
  （2026-09-11 の測定: 法人格の無い名前 9 冊がこれで束の疑いに入れなかった ── 穴の値段）

★ 意図して入れていない疑い:
  「番号がとぶ」── 取引先の請求番号は**その取引先の全顧客**で通しなので、買い手の手元で
  欠番があるのは普通。鳴きすぎる見込みなので入れない（検体では 1 件取り逃す。測って決める）。
"""
from __future__ import annotations

import datetime as dt
import re
import statistics
import unicodedata
from itertools import combinations

#: 所見の種類（★ 語はここにしか書かない）。
DUPLICATE = "重複"
REISSUE = "訂正再発行"
YEAR_OFF = "年の誤り"
MAGNITUDE = "桁違い"
NUMBER_CLASH = "番号が重なる"
DOUBLE_PAY = "二重払い"
BLANK_DATE = "空欄"
ZERO_AMOUNT = "金額が 0"

KINDS = (DUPLICATE, DOUBLE_PAY, REISSUE, YEAR_OFF, MAGNITUDE, NUMBER_CLASH, BLANK_DATE, ZERO_AMOUNT)

#: 桁違いと呼ぶ比（他の月の中央値に対して）。検体の普通の変動は最大 2.3 倍（2026-09-11）。
MAGNITUDE_RATIO = 4.0

#: 年の誤りと呼ぶ、他の冊からの月の距離。年をまたぐ普通の並び（12 月→1 月）は 1。
YEAR_OFF_MONTHS = 6


def _norm(s) -> str:
    return "".join(str(s).split()) if s is not None else ""


#: ハイフンに見える文字（★ 実物の請求番号に混ざる）。`ー` は**長音符** ── かな入力のまま
#: `-` を打つ、実務で一番多い打ち間違い。
_HYPHEN_LIKE = re.compile(r"[ー—–‐‑‒–—―−－~〜]")


def number_key(value) -> str:
    """請求番号を**照合するときだけ**の形に畳む（★ 値は原本のまま・畳んだ形は出さない）。

    ★★ なぜ要るか（2026-09-13・Namakoo 提供の検体で実測）: 実物の請求番号が
      `ＩＮＶー０９ー９９９`（全角英数 ＋ 長音符）だった。同じ請求書が PDF と Excel で 2 通
      来て、片方が `INV-09-999` と書かれていると、**重複が 1 件も鳴らなかった**
      （書き方を揃えると鳴る）。二重払いは取り逃しの中で一番高くつく。
    ★ `form_read.norm` は畳まない ── あちらは NFKC を閉じた文字クラスにだけ当てる設計で、
      **全角は「文字で入っていた」証拠**なので消さない。だから照合用の鍵をここに 1 本置く。
    ★ 畳んで一致しても**原本が違えば必ず名指しする**（`SPELLING_NOTE`）── 併合はしない。
      `split` の `matching_core` と同じ線。
    """
    text = unicodedata.normalize("NFKC", str(value if value is not None else ""))
    return _HYPHEN_LIKE.sub("-", "".join(text.split())).upper()


#: 畳んだ形は同じだが原本が違うときに添える 1 行（人が「同じ番号の別の書き方」と分かるように）。
SPELLING_NOTE = "（★ 請求番号の書き方が違います: {a} / {b} ── 全角・半角や長音符の違い）"


def _number_note(na, nb) -> str:
    """原本が違うときだけ 1 行返す（同じなら空 ── 余計な文を足さない）。"""
    return SPELLING_NOTE.format(a=na, b=nb) if str(na) != str(nb) else ""


def _date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def _month_index(d: dt.date) -> int:
    return d.year * 12 + d.month


def _yen(v) -> str:
    try:
        return f"{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return str(v)


def _ym(d: dt.date) -> str:
    return f"{d.year}年{d.month}月"


def _groups_by_vendor(books: dict) -> dict:
    out: dict = {}
    for name, f in books.items():
        v = _norm(f.get("請求元"))
        if v:
            out.setdefault(v, []).append(name)
    return out


#: 差し替えの徴候になるファイル名の語（★ 実務の言い回し）。
RESEND_WORDS = ("再発行", "差替", "差し替え", "訂正", "修正", "取消", "取り消し")

#: 二重払いを疑う日付の窓（日）。★ 月額固定の請求（毎月同額）で鳴らないための線。
DOUBLE_PAY_DAYS = 7


def _resend_word(*names) -> str:
    """ファイル名に差し替えの語が在れば、その語を返す。"""
    for name in names:
        for word in RESEND_WORDS:
            if word in str(name):
                return word
    return ""


def _is_branch(a: str, b: str) -> bool:
    """b が a の枝番（a + '-2' など）か。"""
    a, b = _norm(a), _norm(b)
    return bool(a) and b.startswith(a) and bool(re.fullmatch(r"[-_(（]?\d+[)）]?", b[len(a):]))


def suspect(books: dict) -> list:
    """束の所見。books: {冊の名: {項目: 値 or None}} → [{"種類", "冊", "理由"}, …]"""
    out: list = []
    paired: set = set()          # 重複／再発行／番号重なり で既に組にした 2 冊

    def emit(kind, names, why):
        out.append({"種類": kind, "冊": list(names), "理由": why})

    # ── 取引先をまたいでも成り立つ疑い: 同じ請求書が 2 通（日付・番号・金額が全部同じ） ──
    for a, b in combinations(sorted(books), 2):
        fa, fb = books[a], books[b]
        if fa.get("請求額") is None or fb.get("請求額") is None:
            continue
        if fa["請求額"] != fb["請求額"]:
            continue
        if _norm(fa.get("請求元")) != _norm(fb.get("請求元")):
            continue
        same_keys, note = [], ""
        for k in ("請求日", "請求番号"):
            va, vb = fa.get(k), fb.get(k)
            if va is None or vb is None:
                continue
            # ★ 請求番号は**畳んだ鍵**で比べる（全角・長音符の違いで重複を取り逃さない）。
            if k == "請求番号":
                if number_key(va) != number_key(vb):
                    break
                note = _number_note(va, vb)
            elif (_date(va) or _norm(va)) != (_date(vb) or _norm(vb)):
                break
            same_keys.append(k)
        else:
            if same_keys:
                paired.add(frozenset((a, b)))
                emit(DUPLICATE, (a, b),
                     f"「{a}」と「{b}」は {'・'.join(same_keys)}・金額（{_yen(fa['請求額'])}）が"
                     f"同じ ── 同じ請求書が 2 通ある（重複）疑い{note}")

    # ── 1 冊で成り立つ疑い: 請求額が 0 ──
    # ★★ 2026-09-13（買い手役の初見・経理）: 明細 0 行・小計/税/合計すべて ¥0 の**白紙の雛形**
    #   （送り間違い）が請求額 `0` の行として一覧に載り、合計にそのまま入った。値は原本のまま
    #   （0 と書いてある）── 黙らせないのはここ。取引先が読めない冊でも鳴らす（1 冊で分かる）。
    for name in sorted(books):
        x = books[name].get("請求額")
        if isinstance(x, (int, float)) and not isinstance(x, bool) and abs(float(x)) < 0.005:
            emit(ZERO_AMOUNT, (name,),
                 f"「{name}」の請求額が 0 です ── 白紙の雛形（送り間違い）か、金額の無い書類の疑い。"
                 "合計に入れる前に確認")

    # ── 取引先ごとの疑い ──
    for vendor, names in _groups_by_vendor(books).items():
        dated = [(n, _date(books[n].get("請求日"))) for n in names]
        dated = [(n, d) for n, d in dated if d]

        # 同じ月に 2 通で金額が違い、かつ**差し替えの徴候**が在る → 訂正再発行
        # ★★ 「同じ月に複数通」だけでは疑わない（2026-09-13・月末の束 125 冊で実測）──
        #   建設・運送・資材では同じ取引先が月に 5 通出すのが商慣行で、C(5,2)=10 組が
        #   一斉に立って**誤報 156 件**になった。疑うのは徴候が在るときだけ。
        for (a, da), (b, db) in combinations(dated, 2):
            if frozenset((a, b)) in paired or (da.year, da.month) != (db.year, db.month):
                continue
            xa, xb = books[a].get("請求額"), books[b].get("請求額")
            if xa is None or xb is None or xa == xb:
                continue
            na, nb = books[a].get("請求番号"), books[b].get("請求番号")
            word = _resend_word(a, b)
            # ★ 理由文の**頭**に徴候を置く ── 良性の行と一字も違わない文にしない。
            if na and nb and number_key(na) == number_key(nb):
                head = f"請求番号が同じ（{na}）{_number_note(na, nb)}"
            elif na and nb and (_is_branch(na, nb) or _is_branch(nb, na)):
                head = f"請求番号が枝番（{na} / {nb}）"
            elif word:
                head = f"ファイル名に「{word}」"
            else:
                continue          # ★ 徴候が無い同月複数請求は商慣行 ── 鳴らさない
            paired.add(frozenset((a, b)))
            emit(REISSUE, (a, b),
                 f"「{a}」と「{b}」は{head} ── 同じ月（{_ym(da)}）の請求で金額が違う"
                 f"（{_yen(xa)} と {_yen(xb)}）── 訂正して再発行された疑い。どちらが有効か確認")

        # 同じ取引先・同じ金額・番号が違う・日付が近い → 二重払い
        # ★★ `重複`（同じ紙が 2 枚）とは別の種類にする ── 買い手の次の一手が違う。
        #   二重払いは**別の紙で同じ金額を 2 回**請求されている形で、取り逃すと一番高くつく。
        # ★ 月額固定の保守料（毎月同額）で鳴らないよう、同じ月か DOUBLE_PAY_DAYS 日以内に限る。
        for (a, da), (b, db) in combinations(dated, 2):
            if frozenset((a, b)) in paired:
                continue
            xa, xb = books[a].get("請求額"), books[b].get("請求額")
            if xa is None or xb is None or xa != xb:
                continue
            na, nb = books[a].get("請求番号"), books[b].get("請求番号")
            if not na or not nb or number_key(na) == number_key(nb):
                continue          # ★ 番号が同じなら別の種類（重複・番号が重なる）の話
            gap = abs((da - db).days)
            same_month = (da.year, da.month) == (db.year, db.month)
            if not same_month and gap > DOUBLE_PAY_DAYS:
                continue
            when = (f"どちらも {_ym(da)} の請求" if same_month
                    else f"請求日が {gap} 日違い（{da.isoformat()} と {db.isoformat()}）")
            paired.add(frozenset((a, b)))
            emit(DOUBLE_PAY, (a, b),
                 f"「{a}」と「{b}」は同額（{_yen(xa)}）で請求番号が違う（{na} / {nb}）・"
                 f"{when} ── 別の紙で同じ金額を 2 回請求されている"
                 f"（二重払い・支払いが重複する）疑い。別の請求か、片方が再請求かを確認")

        # 同じ請求番号で日付か金額が違う → 番号が重なる
        for a, b in combinations(names, 2):
            if frozenset((a, b)) in paired:
                continue
            na, nb = books[a].get("請求番号"), books[b].get("請求番号")
            if not na or not nb or number_key(na) != number_key(nb):
                continue
            paired.add(frozenset((a, b)))
            emit(NUMBER_CLASH, (a, b),
                 f"「{a}」と「{b}」は請求番号（{na}）が重なる（同じ番号が 2 通・重複した番号）"
                 f"{_number_note(na, nb)}"
                 f"のに内容が違う ── 番号の付け間違いか、片方が別の請求か確認")

        # 1 冊だけ年が離れている → 年の誤り
        if len(dated) >= 3:
            for n, d in dated:
                others = [_month_index(e) for m, e in dated if m != n]
                if min(abs(_month_index(d) - o) for o in others) < YEAR_OFF_MONTHS:
                    continue
                if max(others) - min(others) >= YEAR_OFF_MONTHS:
                    continue
                years = sorted({e.year for m, e in dated if m != n})
                emit(YEAR_OFF, (n,),
                     f"「{n}」の請求日（{d.isoformat()}）だけ年が離れている"
                     f"（他の冊は {'・'.join(str(y) for y in years)} 年）── {d.year} 年は年の誤りの疑い")

        # 1 冊だけ桁が違う → 桁違い
        amounts = [(n, books[n].get("請求額")) for n in names]
        amounts = [(n, float(x)) for n, x in amounts if isinstance(x, (int, float)) and x > 0]
        if len(amounts) >= 3:
            for n, x in amounts:
                med = statistics.median(y for m, y in amounts if m != n)
                if med <= 0:
                    continue
                ratio = x / med if x > med else med / x
                if ratio < MAGNITUDE_RATIO or len(str(int(x))) == len(str(int(med))):
                    continue
                side = "多い" if x > med else "少ない"
                emit(MAGNITUDE, (n,),
                     f"「{n}」の金額（{_yen(x)}）は他の月（中央値 {_yen(med)}）の約 {ratio:.0f} 倍"
                     f" ── 桁が 1 つ{side}（10 倍の違い）疑い。0 の数を確認")

        # 他の月には請求日があるのに、この冊だけ空 → 空欄
        if len(names) >= 2 and dated:
            for n in names:
                if books[n].get("請求日") is None:
                    emit(BLANK_DATE, (n,),
                         f"「{n}」は請求日が空（同じ取引先の他の冊 {len(dated)} 冊には請求日がある）"
                         f" ── 請求日の記入漏れの疑い")
    return out
