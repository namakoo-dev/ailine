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
from itertools import combinations

#: 所見の種類（★ 語はここにしか書かない）。
DUPLICATE = "重複"
REISSUE = "訂正再発行"
YEAR_OFF = "年の誤り"
MAGNITUDE = "桁違い"
NUMBER_CLASH = "番号が重なる"
BLANK_DATE = "空欄"

KINDS = (DUPLICATE, REISSUE, YEAR_OFF, MAGNITUDE, NUMBER_CLASH, BLANK_DATE)

#: 桁違いと呼ぶ比（他の月の中央値に対して）。検体の普通の変動は最大 2.3 倍（2026-09-11）。
MAGNITUDE_RATIO = 4.0

#: 年の誤りと呼ぶ、他の冊からの月の距離。年をまたぐ普通の並び（12 月→1 月）は 1。
YEAR_OFF_MONTHS = 6


def _norm(s) -> str:
    return "".join(str(s).split()) if s is not None else ""


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
        same_keys = []
        for k in ("請求日", "請求番号"):
            va, vb = fa.get(k), fb.get(k)
            if va is None or vb is None:
                continue
            if (_date(va) or _norm(va)) != (_date(vb) or _norm(vb)):
                break
            same_keys.append(k)
        else:
            if same_keys:
                paired.add(frozenset((a, b)))
                emit(DUPLICATE, (a, b),
                     f"「{a}」と「{b}」は {'・'.join(same_keys)}・金額（{_yen(fa['請求額'])}）が"
                     f"同じ ── 同じ請求書が 2 通ある（重複）疑い")

    # ── 取引先ごとの疑い ──
    for vendor, names in _groups_by_vendor(books).items():
        dated = [(n, _date(books[n].get("請求日"))) for n in names]
        dated = [(n, d) for n, d in dated if d]

        # 同じ月に 2 通、金額が違う → 訂正再発行
        for (a, da), (b, db) in combinations(dated, 2):
            if frozenset((a, b)) in paired or (da.year, da.month) != (db.year, db.month):
                continue
            xa, xb = books[a].get("請求額"), books[b].get("請求額")
            if xa is None or xb is None or xa == xb:
                continue
            na, nb = books[a].get("請求番号"), books[b].get("請求番号")
            hint = ""
            if na and nb and (_is_branch(na, nb) or _is_branch(nb, na)):
                hint = f"・請求番号に枝番（{na} / {nb}）"
            paired.add(frozenset((a, b)))
            emit(REISSUE, (a, b),
                 f"「{a}」と「{b}」は同じ月（{_ym(da)}）の請求で金額が違う"
                 f"（{_yen(xa)} と {_yen(xb)}）{hint} ── 訂正して再発行された疑い。どちらが有効か確認")

        # 同じ請求番号で日付か金額が違う → 番号が重なる
        for a, b in combinations(names, 2):
            if frozenset((a, b)) in paired:
                continue
            na, nb = books[a].get("請求番号"), books[b].get("請求番号")
            if not na or not nb or _norm(na) != _norm(nb):
                continue
            paired.add(frozenset((a, b)))
            emit(NUMBER_CLASH, (a, b),
                 f"「{a}」と「{b}」は請求番号（{na}）が重なる（同じ番号が 2 通・重複した番号）"
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
